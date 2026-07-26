<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Mail;

use Mnb\PHPExcel\Support\MnbExcelException;

/** Native SMTP transport with STARTTLS/SMTPS, AUTH LOGIN/PLAIN and attachments. */
final class SmtpMailer implements MailerInterface
{
    /** @param array<string,mixed> $options */
    public function __construct(private readonly array $options)
    {
        if(trim((string)($options['host']??''))==='')throw new MnbExcelException('SMTP host is required.');
        if(trim((string)($options['from']??''))==='')throw new MnbExcelException('SMTP from address is required.');
    }

    public function send(MailMessage $message):bool
    {
        $host=(string)$this->options['host'];$encryption=strtolower((string)($this->options['encryption']??'starttls'));$port=(int)($this->options['port']??($encryption==='smtps'?465:587));$timeout=max(1,(int)($this->options['timeout_seconds']??20));
        $context=stream_context_create(['ssl'=>['verify_peer'=>(bool)($this->options['verify_peer']??true),'verify_peer_name'=>(bool)($this->options['verify_peer_name']??true),'allow_self_signed'=>(bool)($this->options['allow_self_signed']??false),'peer_name'=>$host]]);
        $scheme=$encryption==='smtps'?'tls':'tcp';$socket=@stream_socket_client($scheme.'://'.$host.':'.$port,$errno,$errstr,$timeout,STREAM_CLIENT_CONNECT,$context);if(!is_resource($socket))throw new MnbExcelException('SMTP connection failed: '.$errstr.' ('.$errno.')');stream_set_timeout($socket,$timeout);
        try{
            $this->expect($socket,[220]);$hello=(string)($this->options['hello']??(gethostname()?:'localhost'));$this->command($socket,'EHLO '.$hello,[250]);
            if($encryption==='starttls'){$this->command($socket,'STARTTLS',[220]);if(!stream_socket_enable_crypto($socket,true,STREAM_CRYPTO_METHOD_TLS_CLIENT))throw new MnbExcelException('Unable to enable SMTP TLS.');$this->command($socket,'EHLO '.$hello,[250]);}
            $user=(string)($this->options['username']??'');if($user!==''){$method=strtoupper((string)($this->options['auth']??'LOGIN'));$pass=(string)($this->options['password']??'');if($method==='PLAIN'){$this->command($socket,'AUTH PLAIN '.base64_encode("\0".$user."\0".$pass),[235]);}else{$this->command($socket,'AUTH LOGIN',[334]);$this->command($socket,base64_encode($user),[334]);$this->command($socket,base64_encode($pass),[235]);}}
            $from=$this->address((string)$this->options['from']);$this->command($socket,'MAIL FROM:<'.$from.'>',[250]);$recipients=$message->to;foreach(['Cc','Bcc'] as $name){if(isset($message->headers[$name]))$recipients=array_merge($recipients,array_map('trim',explode(',',$message->headers[$name])));}foreach(array_unique($recipients) as $recipient){$this->command($socket,'RCPT TO:<'.$this->address((string)$recipient).'>',[250,251]);}
            $this->command($socket,'DATA',[354]);$mime=(new MimeMessageBuilder())->build($message,(string)$this->options['from']);$raw=preg_replace('/(?m)^\./','..',$mime['raw'])??$mime['raw'];fwrite($socket,$raw."\r\n.\r\n");$this->expect($socket,[250]);$this->command($socket,'QUIT',[221]);return true;
        }finally{fclose($socket);}
    }

    /** @param resource $socket @param list<int> $codes */
    private function command($socket,string $command,array $codes):string{if(fwrite($socket,$command."\r\n")===false)throw new MnbExcelException('Unable to write SMTP command.');return $this->expect($socket,$codes);}
    /** @param resource $socket @param list<int> $codes */
    private function expect($socket,array $codes):string{$response='';do{$line=fgets($socket,8192);if($line===false)throw new MnbExcelException('SMTP server closed the connection.');$response.=$line;$more=isset($line[3])&&$line[3]==='-';}while($more);$code=(int)substr($response,0,3);if(!in_array($code,$codes,true))throw new MnbExcelException('SMTP error '.$code.': '.trim($response));return $response;}
    private function address(string $value):string{if(preg_match('/<([^>]+)>/',$value,$m)===1)$value=$m[1];$value=trim($value);if(filter_var($value,FILTER_VALIDATE_EMAIL)===false)throw new MnbExcelException('Invalid email address: '.$value);return $value;}
}
