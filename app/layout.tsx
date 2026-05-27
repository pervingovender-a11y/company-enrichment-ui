import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Business Ready Finance",
  description: "Funding eligibility calculator preview",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body className="antialiased">{children}</body>
    </html>
  );
}
