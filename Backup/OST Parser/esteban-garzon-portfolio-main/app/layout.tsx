import type { Metadata } from "next";
import { Bricolage_Grotesque, Manrope } from "next/font/google";
import { Analytics } from "@vercel/analytics/react";
import "./globals.css";

const display = Bricolage_Grotesque({
  subsets: ["latin"],
  variable: "--font-display",
});

const body = Manrope({
  subsets: ["latin"],
  variable: "--font-body",
});

export const metadata: Metadata = {
  title: "Esteban Garzon | Automation Engineer",
  description:
    "Esteban Garzon - CTO focused on automation, data engineering, ETL pipelines, AI-driven systems, API integrations, telemetry workflows, Microsoft Graph, RFMS, Moraware, and modern operational software.",
  keywords: [
    "Esteban Garzon",
    "CTO",
    "Automation Engineer",
    "Data Engineer",
    "Business Intelligence",
    "ETL",
    "AI Systems",
    "Python",
    "Portfolio",
    "Outlook Attachment Extractor",
    "Code Lab",
    "Moraware",
    "RFMS",
    "Microsoft Graph",
    "Samsara",
    "Excel Automation",
  ],
  authors: [{ name: "Esteban Garzon" }],
  creator: "Esteban Garzon",
  metadataBase: new URL("https://www.estebangarzon.com"),
  alternates: {
    canonical: "/",
  },
  openGraph: {
    title: "Esteban Garzon | CTO - Data, Automation, and AI Systems",
    description:
      "Portfolio, resume snapshot, technical projects, Code Lab samples, and downloadable tools built by Esteban Garzon.",
    url: "https://www.estebangarzon.com",
    siteName: "Esteban Garzon",
    locale: "en_US",
    type: "website",
  },
  twitter: {
    card: "summary_large_image",
    title: "Esteban Garzon | CTO - Data, Automation, and AI Systems",
    description:
      "Portfolio, resume snapshot, technical projects, Code Lab samples, and downloadable tools built by Esteban Garzon.",
  },
  robots: {
    index: true,
    follow: true,
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body className={`${display.variable} ${body.variable}`}>
        {children}
        <Analytics />
      </body>
    </html>
  );
}
