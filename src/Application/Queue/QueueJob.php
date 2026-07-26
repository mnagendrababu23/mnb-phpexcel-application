<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Queue;

final class QueueJob
{
    /** @param array<string,mixed> $payload */
    public function __construct(
        public readonly string $id,
        public readonly string $type,
        public readonly array $payload,
        public readonly int $attempts = 0,
        public readonly int $availableAt = 0,
        public readonly string $createdAt = ''
    ) {
    }

    /** @param array<string,mixed> $data */
    public static function fromArray(array $data): self
    {
        return new self(
            (string) ($data['id'] ?? ''),
            (string) ($data['type'] ?? ''),
            (array) ($data['payload'] ?? []),
            (int) ($data['attempts'] ?? 0),
            (int) ($data['available_at'] ?? 0),
            (string) ($data['created_at'] ?? '')
        );
    }

    /** @return array<string,mixed> */
    public function toArray(): array
    {
        return [
            'id' => $this->id,
            'type' => $this->type,
            'payload' => $this->payload,
            'attempts' => $this->attempts,
            'available_at' => $this->availableAt,
            'created_at' => $this->createdAt,
        ];
    }
}
