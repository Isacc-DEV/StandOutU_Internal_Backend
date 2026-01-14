export type SourceUrlRecord = {
  id: number;
  url: string;
  sourceName: string;
  countryId: number | null;
};

export type ScrapeStore = {
  listSourceUrls: () => Promise<SourceUrlRecord[]>;
  createJobLink: (url: string, countryId: number | null) => Promise<boolean>;
  close: () => Promise<void>;
};
