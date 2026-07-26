<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Mail;

use Mnb\PHPExcel\Support\MnbExcelException;

final class MailMessage
{
    /** @param list<string> $to @param list<array{path:string,name?:string,mime?:string}> $attachments @param array<string,string> $headers */
    public function __construct(
        public readonly array $to,
        public readonly string $subject,
        public readonly string $body,
        public readonly bool $html = false,
        public readonly array $attachments = [],
        public readonly array $headers = []
    ) {
        if ($to === []) {
            throw new MnbExcelException('At least one email recipient is required.');
        }
        foreach ($to as $address) {
            if (filter_var($address, FILTER_VALIDATE_EMAIL) === false) {
                throw new MnbExcelException('Invalid email recipient: ' . $address);
            }
        }
        foreach ($attachments as $attachment) {
            if (!is_file((string) ($attachment['path'] ?? ''))) {
                throw new MnbExcelException('Email attachment not found: ' . (string) ($attachment['path'] ?? ''));
            }
        }
    }
}
