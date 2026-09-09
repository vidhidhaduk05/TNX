/**
 * tiff-export.js - Standard Baseline uncompressed TIFF generator from Canvas / ImageData
 * Exports publication-quality 300/600 DPI RGB TIFF files.
 */
(function(global) {
  function canvasToTiffBlob(canvas, dpi) {
    dpi = dpi || 300;
    const width = canvas.width;
    const height = canvas.height;
    const ctx = canvas.getContext('2d');
    const imgData = ctx.getImageData(0, 0, width, height);
    const rgba = imgData.data;
    const pixelCount = width * height;
    
    // Convert RGBA -> RGB (with white background blending for any transparent pixels)
    const rgb = new Uint8Array(pixelCount * 3);
    for (let i = 0; i < pixelCount; i++) {
      const a = rgba[i * 4 + 3] / 255;
      if (a === 1) {
        rgb[i * 3]     = rgba[i * 4];
        rgb[i * 3 + 1] = rgba[i * 4 + 1];
        rgb[i * 3 + 2] = rgba[i * 4 + 2];
      } else {
        // Blend over pure white
        rgb[i * 3]     = Math.round(rgba[i * 4] * a + 255 * (1 - a));
        rgb[i * 3 + 1] = Math.round(rgba[i * 4 + 1] * a + 255 * (1 - a));
        rgb[i * 3 + 2] = Math.round(rgba[i * 4 + 2] * a + 255 * (1 - a));
      }
    }

    const offsetIfd = 8;
    const offsetBitsPerSample = 158;
    const offsetRes = 164;
    const offsetStrip = 180;
    const stripByteCount = rgb.length;
    const totalSize = offsetStrip + stripByteCount;
    const buf = new ArrayBuffer(totalSize);
    const view = new DataView(buf);
    const u8 = new Uint8Array(buf);

    // Header: Little Endian ('II'), 42, offset to IFD = 8
    u8[0] = 0x49; u8[1] = 0x49;
    view.setUint16(2, 42, true);
    view.setUint32(4, offsetIfd, true);

    // 12 IFD Tags
    const tags = [
      [256, 4, 1, width],                  // ImageWidth (LONG)
      [257, 4, 1, height],                 // ImageLength (LONG)
      [258, 3, 3, offsetBitsPerSample],    // BitsPerSample (3 * SHORT: 8,8,8)
      [259, 3, 1, 1],                      // Compression = 1 (No compression)
      [262, 3, 1, 2],                      // PhotometricInterpretation = 2 (RGB)
      [273, 4, 1, offsetStrip],            // StripOffsets (LONG)
      [277, 3, 1, 3],                      // SamplesPerPixel = 3 (SHORT)
      [278, 4, 1, height],                 // RowsPerStrip = height (LONG)
      [279, 4, 1, stripByteCount],         // StripByteCounts (LONG)
      [282, 5, 1, offsetRes],              // XResolution (RATIONAL)
      [283, 5, 1, offsetRes + 8],          // YResolution (RATIONAL)
      [296, 3, 1, 2]                       // ResolutionUnit = 2 (Inch)
    ];

    view.setUint16(offsetIfd, tags.length, true);
    let p = offsetIfd + 2;
    for (let i = 0; i < tags.length; i++) {
      const t = tags[i];
      view.setUint16(p, t[0], true);
      view.setUint16(p + 2, t[1], true);
      view.setUint32(p + 4, t[2], true);
      if (t[1] === 3 && t[2] === 1) {
        view.setUint16(p + 8, t[3], true);
        view.setUint16(p + 10, 0, true);
      } else {
        view.setUint32(p + 8, t[3], true);
      }
      p += 12;
    }
    view.setUint32(p, 0, true); // Next IFD = 0

    // BitsPerSample values: 8, 8, 8 (SHORT)
    view.setUint16(offsetBitsPerSample, 8, true);
    view.setUint16(offsetBitsPerSample + 2, 8, true);
    view.setUint16(offsetBitsPerSample + 4, 8, true);

    // XResolution: dpi / 1 (RATIONAL)
    view.setUint32(offsetRes, dpi, true);
    view.setUint32(offsetRes + 4, 1, true);

    // YResolution: dpi / 1 (RATIONAL)
    view.setUint32(offsetRes + 8, dpi, true);
    view.setUint32(offsetRes + 12, 1, true);

    // Image payload
    u8.set(rgb, offsetStrip);

    return new Blob([buf], { type: 'image/tiff' });
  }

  function downloadTiff(canvas, filename, dpi) {
    const blob = canvasToTiffBlob(canvas, dpi || 300);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename.endsWith('.tiff') || filename.endsWith('.tif') ? filename : filename + '.tiff';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 500);
  }

  global.canvasToTiffBlob = canvasToTiffBlob;
  global.downloadTiff = downloadTiff;
})(window);
