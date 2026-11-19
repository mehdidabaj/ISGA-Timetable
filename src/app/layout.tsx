import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "ISGA — Emplois du temps",
  description:
    "Consultation en temps réel des emplois du temps des filières ISGA depuis SharePoint.",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="fr">
      <head>
        <link rel="preconnect" href="https://fonts.googleapis.com" />
        <link rel="preconnect" href="https://fonts.gstatic.com" crossOrigin="" />
        <link
          href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,550;9..144,700&family=Manrope:wght@400;500;650;700&display=swap"
          rel="stylesheet"
        />
      </head>
      <body>{children}</body>
    </html>
  );
}
