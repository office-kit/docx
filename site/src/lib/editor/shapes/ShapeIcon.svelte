<script lang="ts">
  /** A gallery thumbnail of a shape preset, drawn from the same VML path the document gets. */
  import { SHAPE_PRESETS, type ShapePreset, type ShapePresetDef } from '@office-kit/docx';
  import { vmlPathPreviewSvg } from '@office-kit/docx-editor';

  type Props = { preset: ShapePreset; size?: number };
  const { preset, size = 18 }: Props = $props();
  // Room for strokes and callout tails that leave the box.
  const PAD = 3;
  const def: ShapePresetDef = $derived(SHAPE_PRESETS[preset]);
  const inner = $derived(size - PAD * 2);
  const body = $derived.by(() => {
    if (def.element === 'oval') return `<ellipse cx="${inner / 2}" cy="${inner / 2}" rx="${inner / 2}" ry="${inner / 2}"/>`;
    if (def.element === 'rect') return `<rect width="${inner}" height="${inner}"/>`;
    if (def.element === 'roundrect') return `<rect width="${inner}" height="${inner}" rx="${inner / 6}"/>`;
    return def.path ? vmlPathPreviewSvg(def.path, inner) : '';
  });
</script>

<svg width={size} height={size} viewBox="{-PAD} {-PAD} {size} {size}" fill={def.filled === false ? 'none' : '#fff'} stroke="currentColor" stroke-width="1" stroke-linejoin="round" overflow="visible" aria-hidden="true">
  <!-- Trusted markup generated from the library's preset table. -->
  {@html body}
  {#if def.endArrow}<path d="M{inner - 4} {inner - 1} L{inner} {inner} L{inner - 1} {inner - 4}" fill="currentColor" />{/if}
  {#if def.startArrow}<path d="M4 1 L0 0 L1 4" fill="currentColor" />{/if}
</svg>
