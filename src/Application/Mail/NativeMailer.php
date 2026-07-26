<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Mail;

final class NativeMailer implements MailerInterface
{
    /** @param callable(string,string,string,string):bool|null $transport */
    public function __construct(private readonly mixed $transport = null)
    {
    }

    public function send(MailMessage $message): bool
    {
        $boundary = '=_MNB_' . bin2hex(random_bytes(12));
        $headers = [
            'MIME-Version' => '1.0',
            'Content-Type' => 'multipart/mixed; boundary="' . $boundary . '"',
        ] + $message->headers;
        $body = '--' . $boundary . "\r\n"
            . 'Content-Type: ' . ($message->html ? 'text/html' : 'text/plain') . '; charset=UTF-8' . "\r\n"
            . 'Content-Transfer-Encoding: 8bit' . "\r\n\r\n"
            . $message->body . "\r\n";

        foreach ($message->attachments as $attachment) {
            $path = (string) $attachment['path'];
            $name = (string) ($attachment['name'] ?? basename($path));
            $mime = (string) ($attachment['mime'] ?? $this->mime($path));
            $body .= '--' . $boundary . "\r\n"
                . 'Content-Type: ' . $mime . '; name="' . addcslashes($name, '"\\') . '"' . "\r\n"
                . 'Content-Disposition: attachment; filename="' . addcslashes($name, '"\\') . '"' . "\r\n"
                . 'Content-Transfer-Encoding: base64' . "\r\n\r\n"
                . chunk_split(base64_encode((string) file_get_contents($path))) . "\r\n";
        }
        $body .= '--' . $boundary . "--\r\n";
        $headerText = implode("\r\n", array_map(static fn (string $key, string $value): string => $key . ': ' . $value, array_keys($headers), array_values($headers)));
        $to = implode(', ', $message->to);
        if (is_callable($this->transport)) {
            return (bool) ($this->transport)($to, $message->subject, $body, $headerText);
        }
        return mail($to, $message->subject, $body, $headerText);
    }

    private function mime(string $path): string
    {
        if (class_exists(\finfo::class)) {
            $value = (new \finfo(FILEINFO_MIME_TYPE))->file($path);
            if (is_string($value) && $value !== '') {
                return $value;
            }
        }
        return 'application/octet-stream';
    }
}
