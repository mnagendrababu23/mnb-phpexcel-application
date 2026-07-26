<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Http;

final class HttpResponse
{
    /** @param array<string,string> $headers */
    public function __construct(public readonly int $status, public readonly string $body = '', public readonly array $headers = []) {}

    /** @param array<string,mixed> $data @param array<string,string> $headers */
    public static function json(array $data, int $status = 200, array $headers = []): self
    {
        return new self($status, json_encode($data, JSON_UNESCAPED_SLASHES|JSON_UNESCAPED_UNICODE|JSON_THROW_ON_ERROR), ['Content-Type'=>'application/json; charset=UTF-8'] + $headers);
    }

    public function emit(): never
    {
        http_response_code($this->status);
        foreach ($this->headers as $name=>$value) { header($name.': '.$value); }
        echo $this->body;
        exit;
    }
}
