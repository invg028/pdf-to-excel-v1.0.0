export const metadata = {
  title: "PDF → Excel Pro (OCR)",
  description: "Convert scanned or digital PDFs to Excel (client-side, OCR fallback)",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="id">
      <body style={{ fontFamily: "ui-sans-serif, system-ui, Arial", background: "#0b0f17", color: "#e6eaf2" }}>
        <div style={{ maxWidth: 980, margin: "40px auto", padding: 24 }}>
          {children}
        </div>
      </body>
    </html>
  );
}
