/** Insert Chart's types and the sample data Word fills a new chart with. */

import type { ChartGrouping, ChartKind, ChartSpec } from "@office-kit/docx";

export type ChartFamily = "column" | "line" | "pie" | "bar" | "area" | "scatter" | "radar";

export type ChartChoiceId =
  | "colClustered"
  | "colStacked"
  | "colPercent"
  | "line"
  | "lineStacked"
  | "lineMarkers"
  | "pie"
  | "doughnut"
  | "barClustered"
  | "barStacked"
  | "barPercent"
  | "area"
  | "areaStacked"
  | "areaPercent"
  | "scatter"
  | "radar"
  | "radarMarkers";

export interface ChartChoice {
  readonly id: ChartChoiceId;
  readonly kind: ChartKind;
  readonly grouping?: ChartGrouping;
  readonly markers?: boolean;
}

/** The Insert Chart gallery: chart families and their subtypes, in Word's order. */
export const CHART_CHOICES: ReadonlyArray<{
  family: ChartFamily;
  choices: readonly ChartChoice[];
}> = [
  {
    family: "column",
    choices: [
      { id: "colClustered", kind: "column", grouping: "clustered" },
      { id: "colStacked", kind: "column", grouping: "stacked" },
      { id: "colPercent", kind: "column", grouping: "percentStacked" },
    ],
  },
  {
    family: "line",
    choices: [
      { id: "line", kind: "line", grouping: "clustered" },
      { id: "lineStacked", kind: "line", grouping: "stacked" },
      { id: "lineMarkers", kind: "line", grouping: "clustered", markers: true },
    ],
  },
  {
    family: "pie",
    choices: [
      { id: "pie", kind: "pie" },
      { id: "doughnut", kind: "doughnut" },
    ],
  },
  {
    family: "bar",
    choices: [
      { id: "barClustered", kind: "bar", grouping: "clustered" },
      { id: "barStacked", kind: "bar", grouping: "stacked" },
      { id: "barPercent", kind: "bar", grouping: "percentStacked" },
    ],
  },
  {
    family: "area",
    choices: [
      { id: "area", kind: "area", grouping: "clustered" },
      { id: "areaStacked", kind: "area", grouping: "stacked" },
      { id: "areaPercent", kind: "area", grouping: "percentStacked" },
    ],
  },
  {
    family: "scatter",
    choices: [{ id: "scatter", kind: "scatter", markers: true }],
  },
  {
    family: "radar",
    choices: [
      { id: "radar", kind: "radar" },
      { id: "radarMarkers", kind: "radar", markers: true },
    ],
  },
];

const SERIES_DATA: ChartSpec["series"] = [
  { name: "Series 1", values: [4.3, 2.5, 3.5, 4.5] },
  { name: "Series 2", values: [2.4, 4.4, 1.8, 2.8] },
  { name: "Series 3", values: [2, 2, 3, 5] },
];

/** A new chart of `choice`'s type with Word's sample data. */
export function sampleChart(choice: ChartChoice): ChartSpec {
  const base = {
    kind: choice.kind,
    ...(choice.grouping && { grouping: choice.grouping }),
    ...(choice.markers !== undefined && { markers: choice.markers }),
    legend: "bottom" as const,
    dataLabels: false,
  };
  if (choice.kind === "pie" || choice.kind === "doughnut") {
    return {
      ...base,
      title: "Sales",
      legend: "bottom",
      categories: ["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"],
      series: [{ name: "Sales", values: [8.2, 3.2, 1.4, 1.2] }],
    };
  }
  if (choice.kind === "scatter") {
    return {
      ...base,
      title: "Y-Values",
      legend: "none",
      categories: ["0.7", "1.8", "2.6"],
      series: [{ name: "Y-Values", values: [2.7, 3.2, 0.8] }],
    };
  }
  return {
    ...base,
    title: "Chart Title",
    categories: ["Category 1", "Category 2", "Category 3", "Category 4"],
    series: SERIES_DATA,
  };
}

/** The same chart as another type (Change Chart Type keeps the data). */
export function withChartType(spec: ChartSpec, choice: ChartChoice): ChartSpec {
  const { grouping: _grouping, markers: _markers, ...rest } = spec;
  return {
    ...rest,
    kind: choice.kind,
    ...(choice.grouping && { grouping: choice.grouping }),
    ...(choice.markers !== undefined && { markers: choice.markers }),
  };
}
