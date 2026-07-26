<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Mail;

interface MailerInterface
{
    public function send(MailMessage $message): bool;
}
