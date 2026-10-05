<script lang="ts" module>
  export type { IconName } from './icons/index.js';
</script>

<script lang="ts">
  import { iconParts, type IconName } from './icons/index.js';

  type Props = { name: IconName; size?: number };
  const { name, size = 20 }: Props = $props();
</script>

<svg
  viewBox="0 0 20 20"
  width={size}
  height={size}
  fill="none"
  stroke="currentColor"
  stroke-width="1.2"
  stroke-linecap="round"
  stroke-linejoin="round"
  aria-hidden="true"
  >{#each iconParts(name) as part, i (i)}<path d={part.d} stroke={part.stroke} fill={part.fill} />{/each}</svg
>
