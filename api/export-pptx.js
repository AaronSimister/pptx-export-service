import PptxGenJs from 'pptxgenjs';

export default async function handler(req, res) {
  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  // Handle OPTIONS preflight
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  // Only POST
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { presentation } = req.body;

    if (!presentation || !presentation.slides || !Array.isArray(presentation.slides)) {
      return res.status(400).json({
        error: 'Invalid payload: expected { presentation: { title, slides: [] } }'
      });
    }

    const { title = 'Presentation', slides = [] } = presentation;

    if (slides.length === 0) {
      return res.status(400).json({ error: 'At least 1 slide required' });
    }

    // Create presentation
    const prs = new PptxGenJs();
    prs.defineLayout({ name: 'LAYOUT1', width: 10, height: 7.5 });

    // Add slides
    slides.forEach((slide, idx) => {
      const layout = prs.addSlide('LAYOUT1');

      // Header bar
      layout.addShape(prs.ShapeType.rect, {
        x: 0,
        y: 0,
        w: 10,
        h: 0.6,
        fill: { color: 'D97706' }
      });

      // Section text
      layout.addText(slide.section || 'Training', {
        x: 0.3,
        y: 0.05,
        w: 9.4,
        h: 0.5,
        fontSize: 12,
        bold: true,
        color: 'FFFFFF'
      });

      // Title
      layout.addText(slide.title || `Slide ${idx + 1}`, {
        x: 0.5,
        y: 0.8,
        w: 9,
        h: 1,
        fontSize: 32,
        bold: true,
        color: '1F2937'
      });

      // Subtitle
      if (slide.subtitle) {
        layout.addText(slide.subtitle, {
          x: 0.5,
          y: 1.8,
          w: 9,
          h: 0.5,
          fontSize: 18,
          color: '4B5563'
        });
      }

      // Body content
      if (slide.body_content) {
        layout.addText(slide.body_content, {
          x: 0.5,
          y: 2.5,
          w: 9,
          h: 4.5,
          fontSize: 14,
          color: '1F2937',
          breakLine: false
        });
      }
    });

    // Generate PPTX buffer
    const buffer = await prs.write({ outputType: 'arraybuffer' });

    // Convert to base64
    const base64 = Buffer.from(buffer).toString('base64');

    // Filename
    const timestamp = new Date().toISOString().split('T')[0];
    const fileName = `${title.replace(/\s+/g, '-')}-${timestamp}.pptx`;

    return res.status(200).json({
      success: true,
      pptx_base64: base64,
      pptx_size_bytes: buffer.byteLength,
      file_name: fileName,
      slide_count: slides.length
    });
  } catch (error) {
    console.error('[export-pptx] Error:', error);
    return res.status(500).json({
      error: error.message,
      stage: 'generatePptx'
    });
  }
}
