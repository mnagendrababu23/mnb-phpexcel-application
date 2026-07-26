<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Mail;

use Mnb\PHPExcel\Core\WorkbookBuilder;

final class SpreadsheetMailer
{
    public function __construct(private readonly MailerInterface $mailer)
    {
    }

    /** @param string|list<string> $to @param array<string,mixed> $options */
    public function send(WorkbookBuilder|string $workbook, string|array $to, string $subject, string $body = '', array $options = []): bool
    {
        $temporary = false;
        if ($workbook instanceof WorkbookBuilder) {
            $path = (string) ($options['path'] ?? (sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'mnb-export-' . bin2hex(random_bytes(6)) . '.xlsx'));
            $workbook->save($path);
            $temporary = !isset($options['path']);
        } else {
            $path = $workbook;
        }

        try {
            $message = new MailMessage(
                is_array($to) ? array_values($to) : [$to],
                $subject,
                $body,
                (bool) ($options['html'] ?? false),
                array_merge([['path' => $path, 'name' => (string) ($options['filename'] ?? basename($path))]], (array) ($options['attachments'] ?? [])),
                (array) ($options['headers'] ?? [])
            );
            return $this->mailer->send($message);
        } finally {
            if ($temporary) {
                @unlink($path);
            }
        }
    }
}
