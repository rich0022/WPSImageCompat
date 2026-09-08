import JSZip from 'jszip';
export const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6jWQAAAAASUVORK5CYII=';
export function cellImage(id = 'ID_A', relationship = 'rId23'): string {
  return `<etc:cellImage><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="${id}"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${relationship}"/></xdr:blipFill></xdr:pic></etc:cellImage>`;
}
export function imagesXml(images = cellImage()): string {
  return `<etc:cellImages xmlns:etc="http://www.wps.cn/officeDocument/2017/etCustomData" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${images}</etc:cellImages>`;
}
export function relationship(target = 'media/image17.png', extra = ''): string {
  return `<Relationship Id="rId23" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${target}" ${extra}/>`;
}
export function relationshipsXml(items = relationship()): string {
  return `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;
}
/** Synthetic OOXML image parts, not a workbook claimed to have been saved by WPS. */
export function fixture(): JSZip {
  const zip = new JSZip();
  zip.file('xl/cellimages.xml', imagesXml());
  zip.file('xl/_rels/cellimages.xml.rels', relationshipsXml());
  zip.file('xl/media/image17.png', PNG, { base64: true });
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/></Types>');
  return zip;
}
