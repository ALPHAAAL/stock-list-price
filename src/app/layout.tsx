import Head from 'next/head';
import './globals.css';

export const metadata = {
  title: 'Stock app',
  description: 'Created by NOT YOU',
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="en">
      <Head>
        <meta http-equiv='cache-control' content='no-cache' />
        <meta http-equiv='expires' content='0' />
        <meta http-equiv='pragma' content='no-cache' />
      </Head>
      <body>{children}</body>
    </html>
  )
}
