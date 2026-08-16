import JSZip from 'jszip';

const DEV = import.meta.env.DEV;

interface OoxmlProp { k: string; v: string }

export class ExportManager {
  static isValidBase64(str: string): boolean {
    if (typeof str !== 'string' || str.length === 0) return false;
    return /^[A-Za-z0-9+/]*={0,2}$/.test(str);
  }

  static parseDataUrl(dataUrl: string): { mimeType: string; data: string; isBase64: boolean } {
    if (typeof dataUrl !== 'string') {
      throw new Error('Export failed: dataUrl must be a string.');
    }
    if (!dataUrl.includes(',')) {
      throw new Error('Export failed: Invalid data URL format (missing comma).');
    }
    const commaIdx = dataUrl.indexOf(',');
    const header = dataUrl.substring(0, commaIdx);
    const body = dataUrl.substring(commaIdx + 1);

    const mimeMatch = header.match(/^data:([^;]+)/);
    const mimeType = mimeMatch ? mimeMatch[1] : 'application/octet-stream';
    const isBase64 = header.includes(';base64');

    if (DEV) {
      console.log('[ExportManager] parseDataUrl:', {
        headerLength: header.length,
        bodyLength: body.length, mimeType, isBase64,
        bodyPreview: body.substring(0, 80),
      });
    }

    if (isBase64) {
      if (!ExportManager.isValidBase64(body)) {
        throw new Error('Export failed: Invalid Base64 input.');
      }
      return { mimeType, data: body, isBase64: true };
    }

    return { mimeType, data: decodeURIComponent(body), isBase64: false };
  }

  static base64ToBytes(base64: string): Uint8Array {
    if (typeof base64 !== 'string' || base64.length === 0) {
      throw new Error('Export failed: Invalid Base64 input.');
    }
    if (!ExportManager.isValidBase64(base64)) {
      throw new Error('Export failed: Invalid Base64 input.');
    }
    if (DEV) console.log('[ExportManager] base64ToBytes length:', base64.length);
    const raw = atob(base64);
    const bytes = new Uint8Array(raw.length);
    for (let i = 0; i < raw.length; i++) {
      bytes[i] = raw.charCodeAt(i);
    }
    return bytes;
  }

  static sanitizeHtmlContent(html: string): string {
    return html
      .replace(/<img[^>]*src="([^"]*?)"[^>]*>/gi, (_match, src) => {
        if (src && !src.startsWith('data:') && !src.startsWith('http://') && !src.startsWith('https://') && !src.startsWith('blob:')) {
          console.warn('[ExportManager] Removing <img> with invalid src:', src.substring(0, 80));
          return '';
        }
        return _match;
      })
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/on\w+="[^"]*"/gi, '')
      .replace(/on\w+='[^']*'/gi, '');
  }

  static async waitForFonts(): Promise<void> {
    try {
      if (document.fonts && document.fonts.ready) {
        await document.fonts.ready;
        if (DEV) console.log('[ExportManager] Fonts ready');
      }
    } catch (err) {
      console.warn('[ExportManager] Font loading check failed:', err);
    }
  }

  static validateSvg(svg: string): void {
    if (!svg || typeof svg !== 'string') {
      throw new Error('Export failed: SVG input is empty or not a string.');
    }
    if (!svg.includes('xmlns="http://www.w3.org/2000/svg"')) {
      throw new Error('Export failed: SVG is missing required xmlns attribute (xmlns="http://www.w3.org/2000/svg").');
    }
    if (svg.includes('<script') || svg.includes('onload=') || svg.includes('onerror=') || svg.includes('onclick=')) {
      throw new Error('Export failed: SVG contains unsafe or unsupported event handler attributes.');
    }
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(svg, 'image/svg+xml');
      const parseError = doc.querySelector('parsererror');
      if (parseError) {
        console.error('[ExportManager] SVG parse error:', parseError.textContent);
        throw new Error('Export failed: SVG contains malformed XML. ' + (parseError.textContent || '').substring(0, 200));
      }
    } catch (err) {
      if (err instanceof Error && err.message.includes('Export failed')) throw err;
      throw new Error('Export failed: Could not parse SVG for validation.');
    }
  }

  static buildCompositePageSvg(
    page: { width: number; height: number; content?: string },
    canvasDataUrl?: string,
    transparentBg = false
  ): string {
    const w = Math.round(page.width);
    const h = Math.round(page.height);
    const rawContent = (page.content && page.content.trim()) ? page.content : '<p><br></p>';
    const contentHtml = ExportManager.sanitizeHtmlContent(rawContent);
    const imageMarkup = canvasDataUrl
      ? `<image href="${canvasDataUrl}" x="0" y="0" width="${w}" height="${h}" preserveAspectRatio="none" />`
      : '';
    const backgroundMarkup = transparentBg ? '' : `<rect width="${w}" height="${h}" fill="#ffffff" />`;

    const svg = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" xmlns:xhtml="http://www.w3.org/1999/xhtml" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">
  ${backgroundMarkup}
  ${imageMarkup}
  <foreignObject x="0" y="0" width="${w}" height="${h}">
    <div xmlns="http://www.w3.org/1999/xhtml" style="width:${w}px;height:${h}px;overflow:hidden;">
      ${contentHtml}
    </div>
  </foreignObject>
</svg>`;

    return svg;
  }

  static rasterizeCanvasTo(
    canvas: HTMLCanvasElement,
    mimeType: 'image/png' | 'image/jpeg',
    multiplier?: number,
    quality?: number
  ): Promise<string> {
    return new Promise((resolve, reject) => {
      try {
        if (!canvas || !canvas.getContext) {
          reject(new Error('Export failed: Invalid canvas element.'));
          return;
        }
        const ctx = canvas.getContext('2d');
        if (!ctx) {
          reject(new Error('Export failed: Could not get canvas 2D context.'));
          return;
        }
        if (DEV) console.log('[ExportManager] rasterizeCanvasTo:', { mimeType, multiplier, quality });
        resolve(canvas.toDataURL(mimeType, quality ?? (mimeType === 'image/jpeg' ? 0.95 : undefined)));
      } catch (err) {
        reject(new Error('Export failed: Canvas rasterization error: ' + (err instanceof Error ? err.message : 'Unknown')));
      }
    });
  }

  static rasterizeSvg(svg: string, mimeType: 'image/png' | 'image/jpeg', options?: { scale?: number; quality?: number }): Promise<string> {
    return new Promise((resolve, reject) => {
      if (!svg || typeof svg !== 'string') {
        reject(new Error('Export failed: Invalid SVG input.'));
        return;
      }

      try {
        ExportManager.validateSvg(svg);
      } catch (err) {
        reject(err);
        return;
      }

      const blob = new Blob([svg], { type: 'image/svg+xml;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const image = new Image();

      let timedOut = false;
      const timeout = setTimeout(() => {
        timedOut = true;
        URL.revokeObjectURL(url);
        reject(new Error('Export failed: SVG rasterization timed out after 15 seconds. The SVG may contain fonts or images that failed to load.'));
      }, 15000);

      image.onload = () => {
        clearTimeout(timeout);
        if (timedOut) return;
        try {
          const scale = Math.max(0.1, options?.scale || 1);
          const quality = options?.quality ?? (mimeType === 'image/jpeg' ? 0.95 : undefined);
          const canvas = document.createElement('canvas');
          canvas.width = Math.max(1, Math.round((image.width || 1) * scale));
          canvas.height = Math.max(1, Math.round((image.height || 1) * scale));
          const ctx = canvas.getContext('2d');
          if (!ctx) {
            URL.revokeObjectURL(url);
            reject(new Error('Export failed: Could not create raster canvas.'));
            return;
          }
          ctx.fillStyle = '#ffffff';
          ctx.fillRect(0, 0, canvas.width, canvas.height);
          ctx.scale(scale, scale);
          ctx.drawImage(image, 0, 0);
          URL.revokeObjectURL(url);
          if (DEV) console.log('[ExportManager] SVG rasterized successfully:', { w: canvas.width, h: canvas.height, scale });
          resolve(canvas.toDataURL(mimeType, quality));
        } catch (err) {
          URL.revokeObjectURL(url);
          reject(new Error('Export failed: SVG rasterization canvas error: ' + (err instanceof Error ? err.message : 'Unknown')));
        }
      };

      image.onerror = (_, __, ___, ____, error) => {
        clearTimeout(timeout);
        URL.revokeObjectURL(url);
        const detail = error instanceof Error ? error.message : 'Could not render SVG (possibly due to unsupported elements or missing fonts/images)';
        reject(new Error('Export failed: SVG rasterization failed. ' + detail));
      };

      image.src = url;
    });
  }

  static async rasterizeSvgWithFallback(
    svg: string,
    mimeType: 'image/png' | 'image/jpeg',
    fallbackCanvas?: () => HTMLCanvasElement | null,
    options?: { scale?: number; quality?: number }
  ): Promise<string> {
    try {
      await ExportManager.waitForFonts();
      return await ExportManager.rasterizeSvg(svg, mimeType, options);
    } catch (svgErr) {
      console.error('[ExportManager] SVG rasterization failed, trying fallback:', svgErr);
      if (fallbackCanvas) {
        try {
          const fc = fallbackCanvas();
          if (fc) {
            const result = await ExportManager.rasterizeCanvasTo(fc, mimeType, options?.scale, options?.quality);
            console.log('[ExportManager] Fallback canvas export succeeded');
            return result;
          }
        } catch (fallbackErr) {
          console.error('[ExportManager] Fallback canvas export also failed:', fallbackErr);
        }
      }
      throw svgErr;
    }
  }

  static generatePdfBlob(pages: Array<{ dataUrl: string; width: number; height: number }>): Blob {
    const encoder = new TextEncoder();
    const chunks: Uint8Array[] = [];
    const objOffsets: number[] = [];

    const write = (str: string) => chunks.push(encoder.encode(str));
    const writeBinary = (data: Uint8Array) => chunks.push(data);
    const offset = () => chunks.reduce((s, c) => s + c.length, 0);

    if (!pages || pages.length === 0) {
      throw new Error('Export failed: No pages to export.');
    }

    const imgs = pages.map((p, i) => {
      if (!p.dataUrl || typeof p.dataUrl !== 'string') {
        throw new Error(`Export failed: Page ${i + 1} has no image data.`);
      }
      let parsed: ReturnType<typeof ExportManager.parseDataUrl>;
      try {
        parsed = ExportManager.parseDataUrl(p.dataUrl);
      } catch {
        throw new Error(`Export failed: Page ${i + 1} has invalid image data.`);
      }
      if (!parsed.isBase64) {
        throw new Error(`Export failed: Page ${i + 1} must be a raster image (PNG/JPEG). SVG data URLs are not supported for PDF embedding.`);
      }
      const bytes = ExportManager.base64ToBytes(parsed.data);
      return { bytes, width: p.width, height: p.height };
    });

    write('%PDF-1.4\n');

    objOffsets[1] = offset();
    write('1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n');

    const pageRefs = imgs.map((_, i) => `${3 + i * 3} 0 R`).join(' ');
    objOffsets[2] = offset();
    write(`2 0 obj\n<< /Type /Pages /Kids [${pageRefs}] /Count ${imgs.length} >>\nendobj\n`);

    let objNum = 3;
    imgs.forEach((img) => {
      const pageObjNum = objNum++;
      const contentObjNum = objNum++;
      const imageObjNum = objNum++;

      objOffsets[pageObjNum] = offset();
      write(`${pageObjNum} 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${img.width} ${img.height}] /Contents ${contentObjNum} 0 R /Resources << /XObject << /Im0 ${imageObjNum} 0 R >> >> >>\nendobj\n`);

      const stream = `q\n${img.width} 0 0 ${img.height} 0 0 cm\n/Im0 Do\nQ\n`;
      objOffsets[contentObjNum] = offset();
      write(`${contentObjNum} 0 obj\n<< /Length ${stream.length} >>\nstream\n${stream}\nendstream\nendobj\n`);

      objOffsets[imageObjNum] = offset();
      write(`${imageObjNum} 0 obj\n<< /Type /XObject /Subtype /Image /Width ${img.width} /Height ${img.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${img.bytes.length} >>\nstream\n`);
      writeBinary(img.bytes);
      write('\nendstream\nendobj\n');
    });

    const xrefOffset = offset();
    const numObjs = objNum;
    let xref = `xref\n0 ${numObjs}\n0000000000 65535 f \n`;
    for (let i = 1; i < numObjs; i++) {
      xref += `${String(objOffsets[i]).padStart(10, '0')} 00000 n \n`;
    }
    write(xref);
    write(`trailer\n<< /Size ${numObjs} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`);

    const total = chunks.reduce((s, c) => s + c.length, 0);
    const result = new Uint8Array(total);
    let pos = 0;
    for (const c of chunks) { result.set(c, pos); pos += c.length; }
    return new Blob([result], { type: 'application/pdf' });
  }

  // ---------- DOCX (real OOXML) generation ----------

  static escapeXml(value: string): string {
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  static cssColorToHex(color: string): string {
    let c = String(color || '').trim().toLowerCase();
    if (!c || c === 'transparent') return 'auto';
    if (c.startsWith('#')) {
      if (c.length === 4) {
        c = '#' + c.slice(1).split('').map(x => x + x).join('');
      }
      return c.replace(/^#/, '').slice(0, 6);
    }
    try {
      const canvas = document.createElement('canvas');
      canvas.width = 1;
      canvas.height = 1;
      const ctx = canvas.getContext('2d');
      if (!ctx) return '000000';
      ctx.fillStyle = '#000000';
      ctx.fillStyle = c;
      return ctx.fillStyle.replace(/^#/, '').slice(0, 6);
    } catch {
      return '000000';
    }
  }

  static mergeOoxmlProps(inherited: OoxmlProp[], own: OoxmlProp[]): OoxmlProp[] {
    const keys = new Set(inherited.map(p => p.k));
    return [...inherited, ...own.filter(p => !keys.has(p.k))];
  }

  static runXml(text: string, props: OoxmlProp[]): string {
    const rPr = props.length ? `<w:rPr>${props.map(p => p.v).join('')}</w:rPr>` : '';
    return `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;
  }

  static inlineXml(node: Node, inherited: OoxmlProp[] = []): string {
    let out = '';
    node.childNodes.forEach((child) => {
      if (child.nodeType === Node.TEXT_NODE) {
        const text = ExportManager.escapeXml((child.nodeValue || '').replace(/\s+/g, ' '));
        if (text) out += ExportManager.runXml(text, inherited);
        return;
      }
      if (child.nodeType !== Node.ELEMENT_NODE) return;
      const el = child as HTMLElement;
      const tag = el.tagName.toLowerCase();
      if (tag === 'br') {
        out += '<w:r><w:br/></w:r>';
        return;
      }
      if (['ul', 'ol', 'p', 'div', 'li', 'table', 'tr', 'td', 'th', 'blockquote'].includes(tag) || (tag.startsWith('h') && tag.length === 2)) {
        out += ExportManager.inlineXml(el, inherited);
        return;
      }
      const own: OoxmlProp[] = [];
      const style = el.getAttribute('style') || '';
      if (el.tagName === 'B' || el.tagName === 'STRONG' || /font-weight\s*:\s*(bold|700|600)/i.test(style)) own.push({ k: 'b', v: '<w:b/>' });
      if (el.tagName === 'I' || el.tagName === 'EM' || /font-style\s*:\s*italic/i.test(style)) own.push({ k: 'i', v: '<w:i/>' });
      if (el.tagName === 'U' || /text-decoration\s*:\s*underline/i.test(style)) own.push({ k: 'u', v: '<w:u w:val="single"/>' });
      if (['S', 'STRIKE', 'DEL'].includes(el.tagName) || /text-decoration\s*:\s*line-through/i.test(style)) own.push({ k: 'strike', v: '<w:strike/>' });
      if (el.tagName === 'SUB') own.push({ k: 'vert', v: '<w:vertAlign w:val="subscript"/>' });
      if (el.tagName === 'SUP') own.push({ k: 'vert', v: '<w:vertAlign w:val="superscript"/>' });
      if (el.tagName === 'SPAN' && /font-variant\s*:\s*small-caps/i.test(style)) own.push({ k: 'caps', v: '<w:caps/>' });
      const fontMatch = style.match(/font-family\s*:\s*([^;]+)/i);
      if (fontMatch) {
        const font = fontMatch[1].trim();
        own.push({ k: 'font', v: `<w:rFonts w:ascii="${ExportManager.escapeXml(font)}" w:hAnsi="${ExportManager.escapeXml(font)}"/>` });
      }
      const sizeMatch = style.match(/font-size\s*:\s*([\d.]+)px/i);
      if (sizeMatch) {
        const half = Math.round(parseFloat(sizeMatch[1]) * 2);
        if (half > 0) own.push({ k: 'sz', v: `<w:sz w:val="${half}"/><w:szCs w:val="${half}"/>` });
      }
      const colorMatch = style.match(/color\s*:\s*(#[0-9a-fA-F]{3,8}|rgba?\([^)]+\)|[a-zA-Z]+)/i);
      if (colorMatch) own.push({ k: 'color', v: `<w:color w:val="${ExportManager.cssColorToHex(colorMatch[1])}"/>` });
      const bgMatch = style.match(/background(?:-color)?\s*:\s*(#[0-9a-fA-F]{3,8}|rgba?\([^)]+\)|[a-zA-Z]+)/i);
      if (bgMatch) {
        const fill = ExportManager.cssColorToHex(bgMatch[1]);
        if (fill !== 'auto') own.push({ k: 'shd', v: `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>` });
      }
      const merged = ExportManager.mergeOoxmlProps(inherited, own);
      out += ExportManager.inlineXml(el, merged);
    });
    return out;
  }

  static paragraphXml(inline: string, style?: string, listType?: 'bullet' | 'number'): string {
    const pPr: string[] = [];
    if (style) pPr.push(`<w:pStyle w:val="${style}"/>`);
    if (listType === 'bullet') pPr.push('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
    else if (listType === 'number') pPr.push('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>');
    const pPrXml = pPr.length ? `<w:pPr>${pPr.join('')}</w:pPr>` : '';
    return `<w:p>${pPrXml}${inline}</w:p>`;
  }

  static htmlToDocxXml(html: string): string {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const paragraphs: string[] = [];
    const visit = (node: Node, listType: 'bullet' | 'number' | null = null): void => {
      node.childNodes.forEach((child) => {
        if (child.nodeType === Node.TEXT_NODE) {
          const text = (child.nodeValue || '').replace(/\s+/g, ' ').trim();
          if (text) paragraphs.push(ExportManager.paragraphXml(ExportManager.runXml(ExportManager.escapeXml(text), [])));
          return;
        }
        if (child.nodeType !== Node.ELEMENT_NODE) return;
        const el = child as HTMLElement;
        const tag = el.tagName.toLowerCase();
        if (tag === 'br') return;
        if (tag === 'li') {
          paragraphs.push(ExportManager.paragraphXml(ExportManager.inlineXml(el), undefined, listType || 'bullet'));
          return;
        }
        if (tag === 'ul') { visit(el, 'bullet'); return; }
        if (tag === 'ol') { visit(el, 'number'); return; }
        if (tag === 'p' || tag === 'div' || tag === 'blockquote' || (tag.startsWith('h') && tag.length === 2)) {
          paragraphs.push(ExportManager.paragraphXml(ExportManager.inlineXml(el), tag.startsWith('h') ? tag : undefined));
          return;
        }
        if (tag === 'table' || tag === 'tr' || tag === 'td' || tag === 'th') { visit(el, listType); return; }
        if (tag === 'img') { paragraphs.push(ExportManager.paragraphXml('<w:r><w:t>[Image]</w:t></w:r>')); return; }
        paragraphs.push(ExportManager.paragraphXml(ExportManager.inlineXml(el)));
      });
    };
    visit(doc.body);
    return paragraphs.join('');
  }

  static async generateDocxBlob(pages: Array<{ content?: string }>, docName: string): Promise<Blob> {
    const zip = new JSZip();
    const bodyXml = pages
      .map((page, i) => {
        const content = (page.content && page.content.trim()) ? page.content : '<p><br></p>';
        let xml = ExportManager.htmlToDocxXml(ExportManager.sanitizeHtmlContent(content));
        if (i < pages.length - 1) {
          xml += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
        }
        return xml;
      })
      .join('');

    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${bodyXml}
  </w:body>
</w:document>`;

    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;

    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

    const corePropsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${ExportManager.escapeXml(docName)}</dc:title>
  <dc:creator>Word Doc</dc:creator>
  <cp:lastModifiedBy>Word Doc</cp:lastModifiedBy>
  <cp:revision>1</cp:revision>
</cp:coreProperties>`;

    const documentRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>`;

    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="h1"><w:name w:val="heading 1"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="72"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="h2"><w:name w:val="heading 2"/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="60"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="h3"><w:name w:val="heading 3"/><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="40"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="h4"><w:name w:val="heading 4"/><w:pPr><w:keepNext/><w:spacing w:before="120" w:after="32"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="h5"><w:name w:val="heading 5"/><w:pPr><w:keepNext/><w:spacing w:before="100" w:after="24"/></w:pPr><w:rPr><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="h6"><w:name w:val="heading 6"/><w:pPr><w:keepNext/><w:spacing w:before="80" w:after="16"/></w:pPr><w:rPr><w:b/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>
</w:styles>`;

    const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

    zip.file('[Content_Types].xml', contentTypesXml);
    zip.file('_rels/.rels', relsXml);
    zip.file('docProps/core.xml', corePropsXml);
    zip.file('word/document.xml', documentXml);
    zip.file('word/_rels/document.xml.rels', documentRelsXml);
    zip.file('word/styles.xml', stylesXml);
    zip.file('word/numbering.xml', numberingXml);

    return zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE',
    });
  }

  static downloadBlob(blob: Blob, filename: string, onSuccess?: () => void): void {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    onSuccess?.();
  }
}
