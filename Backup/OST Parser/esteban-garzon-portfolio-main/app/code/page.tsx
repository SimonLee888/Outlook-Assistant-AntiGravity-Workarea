import type { Metadata } from "next";
import { CodeLab } from "@/components/code-lab";

export const metadata: Metadata = {
  title: "Code Lab | Esteban Garzon",
  description:
    "Curated code samples, integration patterns, and sanitized public source bundles covering ETL, APIs, telemetry pipelines, Microsoft Graph, RFMS, Moraware, and Excel automation.",
  alternates: {
    canonical: "/code",
  },
};

export default function CodePage() {
  return <CodeLab />;
}
