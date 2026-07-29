import React from 'react';
import { Badge, Tooltip } from '@fluentui/react-components';
import type { UsageCost } from '@autooffice/shared';
import { useTranslation } from '../i18n/index.ts';

function formatTokens(tokens: number): string {
  if (tokens < 1_000) return String(tokens);
  if (tokens < 1_000_000) return `${(tokens / 1_000).toFixed(tokens < 10_000 ? 1 : 0)}K`;
  return `${(tokens / 1_000_000).toFixed(2)}M`;
}

function formatUsd(usd: number): string {
  return `$${usd.toFixed(usd >= 1 ? 2 : 4)}`;
}

export function CostBadge({ cost }: { cost: UsageCost | null | undefined }) {
  const { t } = useTranslation();
  if (!cost || cost.totalTokens === 0) return null;
  const label =
    cost.source === 'estimated' && cost.totalUsd !== null
      ? formatUsd(cost.totalUsd)
      : `${formatTokens(cost.totalTokens)} ${t('cost.tokensShort')}`;
  const source =
    cost.source === 'local-free'
      ? t('cost.localFree')
      : cost.source === 'tokens-only'
        ? t('cost.tokensOnly')
        : t('cost.estimated');
  return (
    <Tooltip
      relationship="description"
      content={t('cost.details', {
        input: formatTokens(cost.inputTokens),
        output: formatTokens(cost.outputTokens),
        source,
      })}
    >
      <Badge appearance="outline" color="informative" size="small">
        {label}
      </Badge>
    </Tooltip>
  );
}
