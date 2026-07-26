<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Mail;

final class CallbackMailer implements MailerInterface
{
    /** @param callable(MailMessage):bool $callback */
    public function __construct(private readonly mixed $callback)
    {
    }

    public function send(MailMessage $message): bool
    {
        return (bool) ($this->callback)($message);
    }
}
