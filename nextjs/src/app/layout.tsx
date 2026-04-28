import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Student Accessible PowerPoint",
  description: "PowerPoint accessibility workflow for student presentations.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
