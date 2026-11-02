import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Negócio em dia — diagnóstico",
  description: "Oito perguntas simples e sugestões práticas para pequenos negócios.",
  icons: {
    icon: "/favicon.svg",
    shortcut: "/favicon.svg",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="pt-BR">
      <body className="antialiased">{children}</body>
    </html>
  );
}
