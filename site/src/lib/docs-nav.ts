// Single source of truth for the docs sidebar and the previous / next pager.

export type DocLink = {
  /** Absolute path on the site, no trailing slash. */
  href: string;
  title: string;
};

export type DocSection = {
  title: string;
  links: DocLink[];
};

export const docSections: DocSection[] = [
  {
    title: "Guides",
    links: [
      { href: "/docs/getting-started", title: "Getting started" },
      { href: "/docs/recipes", title: "Recipes" },
    ],
  },
];

export const allDocLinks: DocLink[] = docSections.flatMap((s) => s.links);
