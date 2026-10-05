// Single source of truth for the docs sidebar and the previous / next pager.

import { defineMessages } from "$lib/i18n/define";
import type { SiteLocale } from "$lib/i18n/locales";

export type DocLink = {
  /** Site-absolute path with no locale prefix and no trailing slash. */
  href: string;
  title: string;
};

export type DocSection = {
  title: string;
  links: DocLink[];
};

const titles = defineMessages({
  en: { guides: "Guides", gettingStarted: "Getting started", recipes: "Recipes" },
  ja: { guides: "ガイド", gettingStarted: "はじめに", recipes: "レシピ" },
});

export function docSections(id: SiteLocale): DocSection[] {
  const t = titles[id];
  return [
    {
      title: t.guides,
      links: [
        { href: "/docs/getting-started", title: t.gettingStarted },
        { href: "/docs/recipes", title: t.recipes },
      ],
    },
  ];
}

export function allDocLinks(id: SiteLocale): DocLink[] {
  return docSections(id).flatMap((s) => s.links);
}
