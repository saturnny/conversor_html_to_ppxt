export type ColorToken = {
  hex: string;
  alpha?: number;
};

export type BorderToken = {
  color: ColorToken;
  widthPx: number;
};

export type HorizontalAlign = "left" | "center" | "right" | "justify";

export type VerticalAlign = "top" | "middle" | "bottom";

export type TextRun = {
  text: string;
  color?: ColorToken | null;
  fontFamily?: string | null;
  fontSizePx?: number | null;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
};

export type ShapeElement = {
  type: "shape";
  kind: "rect" | "roundRect";
  x: number;
  y: number;
  w: number;
  h: number;
  radiusPx?: number;
  fill?: ColorToken | null;
  border?: BorderToken | null;
  order: number;
  zIndex: number;
};

export type TextElement = {
  type: "text";
  x: number;
  y: number;
  w: number;
  h: number;
  text: string;
  runs: TextRun[];
  fontFamily?: string | null;
  fontSizePx: number;
  color?: ColorToken | null;
  align: HorizontalAlign;
  verticalAlign: VerticalAlign;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  order: number;
  zIndex: number;
};

export type ImageElement = {
  type: "image";
  x: number;
  y: number;
  w: number;
  h: number;
  src: string;
  alt?: string;
  order: number;
  zIndex: number;
};

export type TableCellManifest = {
  text: string;
  background?: ColorToken | null;
  color?: ColorToken | null;
  fontFamily?: string | null;
  fontSizePx?: number | null;
  align: HorizontalAlign;
  bold?: boolean;
};

export type TableElement = {
  type: "table";
  x: number;
  y: number;
  w: number;
  h: number;
  colWidths: number[];
  rowHeights: number[];
  rows: TableCellManifest[][];
  border?: BorderToken | null;
  fill?: ColorToken | null;
  order: number;
  zIndex: number;
};

export type SlideElement = ShapeElement | TextElement | ImageElement | TableElement;

export type SlideManifest = {
  id: string;
  title: string;
  width: number;
  height: number;
  background?: ColorToken | null;
  elements: SlideElement[];
  warnings: string[];
};

export type ConversionManifest = {
  fileName: string;
  sourceName: string;
  slides: SlideManifest[];
  warnings: string[];
};

export type RenderedSlide = {
  title: string;
  width: number;
  height: number;
  imageDataUrl: string;
};

export type RenderedPresentation = {
  fileName: string;
  sourceName: string;
  slides: RenderedSlide[];
  warnings: string[];
};
