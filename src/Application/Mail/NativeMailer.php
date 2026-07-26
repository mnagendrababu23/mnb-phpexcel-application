<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Mail;

final class NativeMailer implements MailerInterface
{
    /** @param callable(string,string,string,string):bool|null $transport */
    public function __construct(private readonly mixed $transport=null, private readonly string $from=''){}
    public function send(MailMessage $message):bool
    {
        $from=$this->from!==''?$this->from:(string)($message->headers['From']??('noreply@'.(gethostname()?:'localhost')));$mime=(new MimeMessageBuilder())->build($message,$from);$headers=$mime['headers'];unset($headers['To'],$headers['Subject']);$headerText=implode("\r\n",array_map(static fn($k,$v)=>$k.': '.$v,array_keys($headers),array_values($headers)));$to=implode(', ',$message->to);if(is_callable($this->transport))return(bool)($this->transport)($to,$message->subject,$mime['body'],$headerText);return mail($to,$message->subject,$mime['body'],$headerText);
    }
}
