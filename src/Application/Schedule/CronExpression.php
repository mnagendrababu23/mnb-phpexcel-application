<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Schedule;

use DateTimeInterface;
use Mnb\PHPExcel\Support\MnbExcelException;

final class CronExpression
{
    /** @var list<string> */
    private array $parts;

    public function __construct(private readonly string $expression)
    {
        $this->parts = preg_split('/\s+/', trim($expression)) ?: [];
        if (count($this->parts) !== 5) {
            throw new MnbExcelException('Cron expression must contain five fields: minute hour day month weekday.');
        }
    }

    public function isDue(DateTimeInterface $time): bool
    {
        $values = [(int) $time->format('i'), (int) $time->format('G'), (int) $time->format('j'), (int) $time->format('n'), (int) $time->format('w')];
        $ranges = [[0, 59], [0, 23], [1, 31], [1, 12], [0, 7]];
        foreach ($this->parts as $index => $field) {
            if (!$this->matches($field, $values[$index], $ranges[$index][0], $ranges[$index][1], $index === 4)) {
                return false;
            }
        }
        return true;
    }

    private function matches(string $field, int $value, int $min, int $max, bool $weekday): bool
    {
        foreach (explode(',', $field) as $segment) {
            $segment = trim($segment);
            $step = 1;
            if (str_contains($segment, '/')) {
                [$segment, $stepText] = explode('/', $segment, 2);
                $step = max(1, (int) $stepText);
            }
            if ($segment === '*') {
                if (($value - $min) % $step === 0) {
                    return true;
                }
                continue;
            }
            if (str_contains($segment, '-')) {
                [$start, $end] = array_map('intval', explode('-', $segment, 2));
                if ($weekday) {
                    $start %= 7;
                    $end %= 7;
                }
                if ($value >= $start && $value <= $end && ($value - $start) % $step === 0) {
                    return true;
                }
                continue;
            }
            $candidate = (int) $segment;
            if ($weekday) {
                $candidate %= 7;
            }
            if ($value === $candidate) {
                return true;
            }
        }
        return false;
    }

    public function __toString(): string
    {
        return $this->expression;
    }
}
