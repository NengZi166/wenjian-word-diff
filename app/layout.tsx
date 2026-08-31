import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: '文鉴｜离线 Word 文档比对',
  description: '在浏览器本地比较两份 Word 文档，识别文字、表格和公式变化。',
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
