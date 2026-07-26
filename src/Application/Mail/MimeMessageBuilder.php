<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Mail;

final class MimeMessageBuilder
{
    /** @return array{headers:array<string,string>,body:string,raw:string} */
    public function build(MailMessage $message, string $from): array
    {
        $boundary='=_MNB_'.bin2hex(random_bytes(12));
        $headers=['From'=>$from,'To'=>implode(', ',$message->to),'Subject'=>$this->encodeHeader($message->subject),'Date'=>gmdate(DATE_RFC2822),'Message-ID'=>'<'.bin2hex(random_bytes(12)).'@'.(gethostname()?:'localhost').'>','MIME-Version'=>'1.0','Content-Type'=>'multipart/mixed; boundary="'.$boundary.'"']+$message->headers;
        unset($headers['Bcc']);
        $body='--'.$boundary."\r\n".'Content-Type: '.($message->html?'text/html':'text/plain').'; charset=UTF-8'."\r\n".'Content-Transfer-Encoding: 8bit'."\r\n\r\n".$message->body."\r\n";
        foreach($message->attachments as $attachment){$path=(string)$attachment['path'];$name=(string)($attachment['name']??basename($path));$mime=(string)($attachment['mime']??$this->mime($path));$bytes=file_get_contents($path);if($bytes===false)throw new \RuntimeException('Unable to read mail attachment: '.$path);$safe=addcslashes($name,'"\\');$body.='--'.$boundary."\r\n".'Content-Type: '.$mime.'; name="'.$safe.'"'."\r\n".'Content-Disposition: attachment; filename="'.$safe.'"'."\r\n".'Content-Transfer-Encoding: base64'."\r\n\r\n".chunk_split(base64_encode($bytes))."\r\n";}
        $body.='--'.$boundary."--\r\n";$headerText=implode("\r\n",array_map(static fn($k,$v)=>$k.': '.$v,array_keys($headers),array_values($headers)));
        return ['headers'=>$headers,'body'=>$body,'raw'=>$headerText."\r\n\r\n".$body];
    }
    private function encodeHeader(string $value):string{return preg_match('/[^\x20-\x7E]/',$value)===1?'=?UTF-8?B?'.base64_encode($value).'?=':$value;}
    private function mime(string $path):string{if(class_exists(\finfo::class)){$v=(new \finfo(FILEINFO_MIME_TYPE))->file($path);if(is_string($v)&&$v!=='')return $v;}return 'application/octet-stream';}
}
