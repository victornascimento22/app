import type { Metadata } from "next";
import localFont from "next/font/local";
import "./globals.css";

const panton = localFont({
  src: './fonts/Panton-Trial-Regular.woff',
  variable: '--font-panton',
});

export const metadata: Metadata = {
  title: "Relatório ",
  description: "Gerador de relatórios ",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="pt-BR">
      <body className={`${panton.variable} font-panton`}>
        {children}
      </body>
    </html>
  );
}
