import { seedScrapeData } from "./index";

seedScrapeData().catch((error) => {
  console.error("Seed failed:", error);
  process.exit(1);
});
