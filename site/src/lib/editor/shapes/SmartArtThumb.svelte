<script lang="ts">
  /** A SmartArt gallery thumbnail: the layout's arrangement in the given colours. */
  import type { SmartArtLayout } from '@office-kit/docx';

  type Props = {
    layout: SmartArtLayout;
    colors?: readonly string[];
    /** SmartArt Styles: outline width and shadow depth of the shapes. */
    lineWidth?: number;
    effect?: number;
  };
  const { layout, colors = ['4472C4'], lineWidth = 0.8, effect = 0 }: Props = $props();
  const c = (i: number): string => `#${colors[i % colors.length] ?? '4472C4'}`;
  const CYCLE_NODES = 5;
  const cycle = Array.from({ length: CYCLE_NODES }, (_, i) => {
    const a = (i / CYCLE_NODES) * 2 * Math.PI - Math.PI / 2;
    return { x: 30 + 17 * Math.cos(a), y: 22 + 15 * Math.sin(a) };
  });
</script>

<svg width="60" height="44" viewBox="0 0 60 44" stroke="#fff" stroke-width={lineWidth} style:filter={effect ? `drop-shadow(0 ${effect}px ${effect}px rgb(0 0 0 / 40%))` : undefined} aria-hidden="true">
  {#if layout === 'basicBlockList'}
    {#each [0, 1, 2, 3, 4, 5] as i (i)}
      <rect x={4 + (i % 3) * 18} y={6 + Math.floor(i / 3) * 17} width="16" height="14" fill={c(i)} />
    {/each}
  {:else if layout === 'verticalBulletList'}
    {#each [0, 1, 2] as i (i)}
      <rect x="6" y={4 + i * 13} width="48" height="10" rx="2" fill={c(i)} />
    {/each}
  {:else if layout === 'basicProcess'}
    {#each [0, 1, 2] as i (i)}
      <rect x={3 + i * 20} y="14" width="14" height="16" fill={c(i)} />
      {#if i < 2}<path d="M{18 + i * 20} 19 l3 3 l-3 3 z" fill="#{colors[0] ?? '4472C4'}" stroke="none" opacity="0.6" />{/if}
    {/each}
  {:else if layout === 'basicCycle'}
    <circle cx="30" cy="22" r="15" fill="none" stroke="#{colors[0] ?? '4472C4'}" stroke-opacity="0.4" />
    {#each cycle as p, i (i)}
      <circle cx={p.x} cy={p.y} r="6" fill={c(i)} />
    {/each}
  {:else}
    <path d="M30 12v6M14 18h32M14 18v6M46 18v6" fill="none" stroke="#{colors[0] ?? '4472C4'}" stroke-opacity="0.6" />
    <rect x="22" y="3" width="16" height="9" fill={c(0)} />
    <rect x="6" y="24" width="16" height="9" fill={c(1)} />
    <rect x="38" y="24" width="16" height="9" fill={c(2)} />
  {/if}
</svg>
