import type { Metadata } from "next";
import { AnalyticsLab } from "@/components/analytics-lab";

export const metadata: Metadata = {
  title: "Applied Analytics | Esteban Garzon",
  description:
    "Forecasting, business intelligence, data visualization, graph analysis, and analytics case work curated as a public-facing portfolio layer for Esteban Garzon.",
  alternates: {
    canonical: "/analytics",
  },
};

export default function AnalyticsPage() {
  return <AnalyticsLab />;
}
