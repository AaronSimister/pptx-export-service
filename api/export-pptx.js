// Vercel Serverless Function: PPTX Export with Phase 2 Visual Fidelity
// Uses pptxgenjs to generate branded slides matching TrainingDeckLayouts
// Node.js 18+, Express middleware compatible

import { Buffer } from 'buffer';
import PptxGenJS from 'pptxgenjs';

// Design tokens (match /lib/deckDesignTokens.js)
const COLORS = {
  copper: '#D97706',
  copperDark: '#B45309',
  charcoal: '#1F2937',
  slate: '#6B7280',
  slateLight: '#F3F4F6',
  white: '#FFFFFF',
  red: '#DC2626',
  redDark: '#991B1B',
  redLight: '#FEE2E2',
  amber: '#F59E0B',
  amberLight: '#FEF3C7',
  blue: '#3B82F6',
  blueDark: '#1E40AF',
};

const TYPOGRAPHY = {
  slideTitle: { size: 32, bold: true, lineSpacing: 24 },
  sectionLabel: { size: 12, bold: true, charSpacing: 10 },
  subtitle: { size: 16, bold: false },
  body: { size: 15, bold: false, lineSpacing: 24 },
  small: { size: 12, bold: false },
  objective: { size: 14, bold: true },
};

// EMU conversions
const EMU = {
  slideWidth: 9144000,
  slideHeight: 6858000,
  marginInch: 0.5,
  marginEmu: 457200,
};

const toEmu = (inches) => inches * 914400;
const toInches = (emu) => emu / 914400;

// ============================================================================
// STANDARD LEARNING SLIDE LAYOUT
// ============================================================================
function addStandardLearningSlide(prs, slide, slideNum, brandColor) {
  const primaryColor = brandColor || COLORS.copper;
  const slid = prs.addSlide();
  
  // Background
  slid.background = { color: COLORS.white };

  // ─── HEADER BAR (44px / 0.458 inches) ───
  slid.addShape(prs.ShapeType.rect, {
    x: 0, y: 0,
    w: '100%', h: 0.458,
    fill: { color: primaryColor },
    line: { type: 'none' }
  });

  // Header accent bar (1px vertical)
  slid.addShape(prs.ShapeType.rect, {
    x: 0.2, y: 0.05,
    w: 0.01, h: 0.358,
    fill: { color: primaryColor },
    line: { type: 'none' }
  });

  // Header text: section → title breadcrumb
  const headerText = `${slide.section || 'Training'} → ${slide.title}`;
  slid.addText(headerText, {
    x: 0.4, y: 0.08,
    w: 8.5, h: 0.3,
    fontFace: 'Calibri',
    fontSize: 11,
    color: COLORS.white,
    bold: false,
    align: 'left',
    valign: 'middle'
  });

  // ─── CONTENT AREA (2 columns: 50% image, 50% text) ───
  const headerBottomY = 0.458;
  const footerHeight = 0.292;
  const contentHeight = 6.25 - headerBottomY - footerHeight; // ~5.5 inches
  const colWidth = 4.25; // 50% of usable width
  const marginH = 0.4;
  const marginV = 0.3;

  const imageX = marginH;
  const imageY = headerBottomY + marginV;
  const imageWidth = colWidth - marginH;
  const imageHeight = contentHeight - marginV * 2;

  const textX = imageX + colWidth;
  const textY = imageY;
  const textWidth = colWidth - marginH;

  // IMAGE COLUMN
  if (slide.generated_image_url && slide.generated_image_url.trim()) {
    try {
      slid.addImage({
        path: slide.generated_image_url,
        x: imageX, y: imageY,
        w: imageWidth, h: imageHeight,
      });
      // Border around image
      slid.addShape(prs.ShapeType.rect, {
        x: imageX, y: imageY,
        w: imageWidth, h: imageHeight,
        fill: { type: 'none' },
        line: { color: COLORS.slateLight, width: 1 }
      });
    } catch (imgErr) {
      // Fallback if image fails
      slid.addShape(prs.ShapeType.rect, {
        x: imageX, y: imageY,
        w: imageWidth, h: imageHeight,
        fill: { color: COLORS.slateLight },
        line: { color: COLORS.slate, width: 1, dashType: 'dash' }
      });
      slid.addText('Image failed to load', {
        x: imageX, y: imageY + imageHeight / 2 - 0.2,
        w: imageWidth, h: 0.4,
        fontSize: 11,
        color: COLORS.slate,
        align: 'center',
        valign: 'middle'
      });
    }
  } else {
    // Placeholder
    slid.addShape(prs.ShapeType.rect, {
      x: imageX, y: imageY,
      w: imageWidth, h: imageHeight,
      fill: { color: COLORS.slateLight },
      line: { type: 'none' }
    });
    slid.addText('Image pending', {
      x: imageX, y: imageY + imageHeight / 2 - 0.15,
      w: imageWidth, h: 0.3,
      fontSize: 12,
      color: COLORS.slate,
      align: 'center',
      valign: 'middle'
    });
  }

  // TEXT COLUMN
  let textCursorY = textY;

  // Title
  slid.addText(slide.title, {
    x: textX, y: textCursorY,
    w: textWidth, h: 0.6,
    fontSize: 32,
    bold: true,
    color: COLORS.charcoal,
    fontFace: 'Calibri',
    lineSpacing: 24,
    wrap: true
  });
  textCursorY += 0.7;

  // Bullets (max 3) — strip leading markers only once
  const bullets = (slide.body_content || '')
    .split('\n')
    .map(l => {
      let text = l.trim();
      // Remove leading number + period (e.g., "1. ")
      text = text.replace(/^[\d]+\.\s+/, '');
      // Remove leading bullet/dash (e.g., "• " or "- ")
      text = text.replace(/^[•\-]\s+/, '');
      return text.trim();
    })
    .filter(l => l && !l.endsWith(':'))
    .slice(0, 3);

  if (bullets.length > 0) {
    bullets.forEach((bullet, idx) => {
      // Number badge
      slid.addText(`${idx + 1}.`, {
        x: textX, y: textCursorY,
        w: 0.25, h: 0.3,
        fontSize: 16,
        bold: true,
        color: primaryColor,
        fontFace: 'Calibri',
        align: 'left'
      });

      // Bullet text (no double-stripping)
      slid.addText(bullet, {
        x: textX + 0.35, y: textCursorY,
        w: textWidth - 0.35, h: 0.8,
        fontSize: 15,
        color: COLORS.charcoal,
        fontFace: 'Calibri',
        wrap: true,
        lineSpacing: 18
      });
      textCursorY += 0.9;
    });
  }

  // Learning outcome callout (if exists)
  if (slide.key_learning_outcomes && slide.key_learning_outcomes.length > 0) {
    textCursorY += 0.2;
    const calloutH = 0.8;
    slid.addShape(prs.ShapeType.rect, {
      x: textX, y: textCursorY,
      w: textWidth, h: calloutH,
      fill: { color: primaryColor, transparency: 90 },
      line: { color: primaryColor, width: 2, dashType: 'solid' }
    });
    slid.addText(slide.key_learning_outcomes[0], {
      x: textX + 0.15, y: textCursorY + 0.1,
      w: textWidth - 0.3, h: calloutH - 0.2,
      fontSize: 14,
      bold: true,
      color: COLORS.charcoal,
      fontFace: 'Calibri',
      wrap: true,
      lineSpacing: 18
    });
  }

  // ─── FOOTER ───
  const footerY = 6.25 - footerHeight;
  slid.addShape(prs.ShapeType.rect, {
    x: 0, y: footerY,
    w: '100%', h: footerHeight,
    fill: { color: COLORS.white },
    line: { color: COLORS.slateLight, width: 1 }
  });

  const footerText = `Slide ${slideNum} | ${slide.section || 'Training'} | ${slide.title}`;
  slid.addText(footerText, {
    x: 0.3, y: footerY + 0.05,
    w: 8, h: 0.2,
    fontSize: 11,
    color: COLORS.slate,
    fontFace: 'Calibri',
    align: 'left',
    valign: 'middle'
  });

  slid.addText('© 2026', {
    x: 8.5, y: footerY + 0.05,
    w: 0.8, h: 0.2,
    fontSize: 11,
    color: COLORS.slateLight,
    fontFace: 'Calibri',
    align: 'right',
    valign: 'middle'
  });
}

// ============================================================================
// SAFETY CRITICAL SLIDE LAYOUT
// ============================================================================
function addSafetyCriticalSlide(prs, slide, slideNum, brandColor) {
  const slid = prs.addSlide();
  slid.background = { color: COLORS.white };

  // ─── SAFETY HEADER (Orange) ───
  slid.addShape(prs.ShapeType.rect, {
    x: 0, y: 0,
    w: '100%', h: 0.458,
    fill: { color: '#FF6B35' },
    line: { type: 'none' }
  });

  const headerText = `⚠️ SAFETY CRITICAL — Mandatory Compliance Required`;
  slid.addText(headerText, {
    x: 0.3, y: 0.1,
    w: 8.5, h: 0.3,
    fontSize: 13,
    bold: true,
    color: COLORS.white,
    fontFace: 'Calibri',
    align: 'left',
    valign: 'middle'
  });

  // ─── CONTENT AREA ───
  const headerBottomY = 0.458;
  const footerHeight = 0.292;
  const contentHeight = 6.25 - headerBottomY - footerHeight;

  const colWidth = 4.25;
  const marginH = 0.4;
  const marginV = 0.3;

  // Layout: image left (if exists), text right
  let textX = marginH;
  let textWidth = 8.2 - marginH * 2;

  if (slide.generated_image_url && slide.generated_image_url.trim()) {
    const imageX = marginH;
    const imageY = headerBottomY + marginV;
    const imageWidth = colWidth - marginH;
    const imageHeight = contentHeight - marginV * 2;

    try {
      slid.addImage({
        path: slide.generated_image_url,
        x: imageX, y: imageY,
        w: imageWidth, h: imageHeight
      });
    } catch (imgErr) {
      // Fallback: show placeholder
      slid.addShape(prs.ShapeType.rect, {
        x: imageX, y: imageY,
        w: imageWidth, h: imageHeight,
        fill: { color: COLORS.slateLight },
        line: { color: COLORS.slate, width: 1, dashType: 'dash' }
      });
    }

    textX = imageX + colWidth;
    textWidth = colWidth - marginH;
  }

  let textCursorY = headerBottomY + marginV;

  // Title
  slid.addText(slide.title, {
    x: textX, y: textCursorY,
    w: textWidth, h: 0.5,
    fontSize: 28,
    bold: true,
    color: COLORS.charcoal,
    fontFace: 'Calibri',
    wrap: true
  });
  textCursorY += 0.6;

  // Bullets (SOP facts) — deduplicated
  const bulletLines = (slide.body_content || '')
    .split('\n')
    .map(l => {
      let text = l.trim();
      text = text.replace(/^[\d]+\.\s+/, '');
      text = text.replace(/^[•\-]\s+/, '');
      return text.trim();
    })
    .filter(l => l)
    .slice(0, 3);

  if (bulletLines.length > 0) {
    bulletLines.forEach((bullet, idx) => {
      slid.addText(`${idx + 1}.`, {
        x: textX, y: textCursorY,
        w: 0.25, h: 0.25,
        fontSize: 14,
        bold: true,
        color: '#FF6B35',
        fontFace: 'Calibri'
      });

      slid.addText(bullet, {
        x: textX + 0.35, y: textCursorY,
        w: textWidth - 0.35, h: 0.7,
        fontSize: 13,
        color: COLORS.charcoal,
        fontFace: 'Calibri',
        wrap: true,
        lineSpacing: 16
      });
      textCursorY += 0.8;
    });
  }

  // Hazards box (if no bullets)
  if (bulletLines.length === 0 && slide.safety_callouts && slide.safety_callouts.length > 0) {
    slid.addShape(prs.ShapeType.rect, {
      x: textX, y: textCursorY,
      w: textWidth, h: 1.2,
      fill: { color: COLORS.redLight },
      line: { color: COLORS.red, width: 2 }
    });

    let hazardY = textCursorY + 0.1;
    slid.addText('Hazards:', {
      x: textX + 0.15, y: hazardY,
      w: textWidth - 0.3, h: 0.25,
      fontSize: 12,
      bold: true,
      color: COLORS.redDark,
      fontFace: 'Calibri'
    });
    hazardY += 0.3;

    slide.safety_callouts.slice(0, 2).forEach((hazard) => {
      slid.addText(`● ${hazard}`, {
        x: textX + 0.15, y: hazardY,
        w: textWidth - 0.3, h: 0.35,
        fontSize: 12,
        color: COLORS.charcoal,
        fontFace: 'Calibri',
        wrap: true
      });
      hazardY += 0.4;
    });
  }

  // Learning outcome callout (if exists) — amber
  if (slide.key_learning_outcomes && slide.key_learning_outcomes.length > 0) {
    textCursorY += 1.5;
    if (textCursorY < 5.5) {
      slid.addShape(prs.ShapeType.rect, {
        x: textX, y: textCursorY,
        w: textWidth, h: 0.7,
        fill: { color: COLORS.amberLight },
        line: { color: COLORS.amber, width: 2 }
      });

      slid.addText(slide.key_learning_outcomes[0], {
        x: textX + 0.15, y: textCursorY + 0.1,
        w: textWidth - 0.3, h: 0.5,
        fontSize: 13,
        bold: true,
        color: COLORS.charcoal,
        fontFace: 'Calibri',
        wrap: true
      });
    }
  }

  // ─── FOOTER ───
  const footerY = 6.25 - footerHeight;
  slid.addShape(prs.ShapeType.rect, {
    x: 0, y: footerY,
    w: '100%', h: footerHeight,
    fill: { color: COLORS.white },
    line: { color: COLORS.slateLight, width: 1 }
  });

  const footerText = `Slide ${slideNum} | ${slide.section || 'Training'} | Safety Critical`;
  slid.addText(footerText, {
    x: 0.3, y: footerY + 0.05,
    w: 8, h: 0.2,
    fontSize: 11,
    color: COLORS.slate,
    fontFace: 'Calibri',
    align: 'left'
  });

  slid.addText('© 2026', {
    x: 8.5, y: footerY + 0.05,
    w: 0.8, h: 0.2,
    fontSize: 11,
    color: COLORS.slateLight,
    fontFace: 'Calibri',
    align: 'right'
  });
}

// ============================================================================
// PROCESS CHECKLIST SLIDE LAYOUT
// ============================================================================
function addProcessChecklistSlide(prs, slide, slideNum, brandColor) {
  const primaryColor = brandColor || COLORS.copper;
  const slid = prs.addSlide();
  slid.background = { color: COLORS.white };

  // Header
  slid.addShape(prs.ShapeType.rect, {
    x: 0, y: 0,
    w: '100%', h: 0.458,
    fill: { color: primaryColor },
    line: { type: 'none' }
  });

  slid.addText(`${slide.section || 'Training'} → ${slide.title}`, {
    x: 0.4, y: 0.1,
    w: 8.5, h: 0.3,
    fontSize: 11,
    bold: false,
    color: COLORS.white,
    fontFace: 'Calibri'
  });

  // Content area: image left (50%), checklist right (50%)
  const headerBottomY = 0.458;
  const footerHeight = 0.292;
  const contentHeight = 6.25 - headerBottomY - footerHeight;
  const colWidth = 4.25;

  const imageX = 0.4;
  const imageY = headerBottomY + 0.3;
  const imageWidth = colWidth - 0.4;
  const imageHeight = contentHeight - 0.6;

  const checklistX = imageX + colWidth;
  const checklistWidth = colWidth - 0.4;

  // Image (if exists)
  if (slide.generated_image_url && slide.generated_image_url.trim()) {
    try {
      slid.addImage({
        path: slide.generated_image_url,
        x: imageX, y: imageY,
        w: imageWidth, h: imageHeight
      });
    } catch (imgErr) {
      // Fallback: show placeholder
      slid.addShape(prs.ShapeType.rect, {
        x: imageX, y: imageY,
        w: imageWidth, h: imageHeight,
        fill: { color: COLORS.slateLight },
        line: { color: COLORS.slate, width: 1, dashType: 'dash' }
      });
    }
  }

  // Checklist title
  let checklistY = imageY;
  slid.addText(slide.title, {
    x: checklistX, y: checklistY,
    w: checklistWidth, h: 0.5,
    fontSize: 30,
    bold: true,
    color: COLORS.charcoal,
    fontFace: 'Calibri',
    wrap: true
  });
  checklistY += 0.6;

  // Checklist items (max 5) — deduplicated
  const checklistItems = (slide.body_content || '')
    .split('\n')
    .map(l => {
      let text = l.trim();
      text = text.replace(/^[\d]+\.\s+/, '');
      text = text.replace(/^[•\-]\s+/, '');
      return text.trim();
    })
    .filter(l => l)
    .slice(0, 5);

  checklistItems.forEach((bullet, idx) => {
    // Checkbox (unchecked)
    slid.addShape(prs.ShapeType.rect, {
      x: checklistX, y: checklistY + 0.05,
      w: 0.2, h: 0.2,
      fill: { type: 'none' },
      line: { color: primaryColor, width: 2 }
    });

    // Bullet text
    slid.addText(bullet, {
      x: checklistX + 0.3, y: checklistY,
      w: checklistWidth - 0.3, h: 0.5,
      fontSize: 13,
      color: COLORS.charcoal,
      fontFace: 'Calibri',
      wrap: true,
      valign: 'top'
    });

    checklistY += 0.65;
  });

  // Quality check callout (if exists)
  if (slide.quality_callouts && slide.quality_callouts[0]) {
    checklistY += 0.2;
    slid.addShape(prs.ShapeType.rect, {
      x: checklistX, y: checklistY,
      w: checklistWidth, h: 0.6,
      fill: { color: COLORS.blue + '15' },
      line: { color: COLORS.blue, width: 2 }
    });

    slid.addText(`✓ ${slide.quality_callouts[0]}`, {
      x: checklistX + 0.15, y: checklistY + 0.1,
      w: checklistWidth - 0.3, h: 0.4,
      fontSize: 12,
      bold: true,
      color: '#0C4A6E',
      fontFace: 'Calibri',
      wrap: true
    });
  }

  // Footer
  const footerY = 6.25 - footerHeight;
  slid.addShape(prs.ShapeType.rect, {
    x: 0, y: footerY,
    w: '100%', h: footerHeight,
    fill: { color: COLORS.white },
    line: { color: COLORS.slateLight, width: 1 }
  });

  slid.addText(`Slide ${slideNum} | ${slide.section || 'Training'} | ${slide.title}`, {
    x: 0.3, y: footerY + 0.05,
    w: 8, h: 0.2,
    fontSize: 11,
    color: COLORS.slate,
    fontFace: 'Calibri'
  });

  slid.addText('© 2026', {
    x: 8.5, y: footerY + 0.05,
    w: 0.8, h: 0.2,
    fontSize: 11,
    color: COLORS.slateLight,
    fontFace: 'Calibri',
    align: 'right'
  });
}

// ============================================================================
// COVER SLIDE LAYOUT
// ============================================================================
function addCoverSlide(prs, slide, brandColor) {
  const primaryColor = brandColor || COLORS.copper;
  const slid = prs.addSlide();
  
  // Full-color background
  slid.background = { color: primaryColor };

  // Accent bar (vertical gradient-like stripe)
  slid.addShape(prs.ShapeType.rect, {
    x: 0, y: 0,
    w: 0.15, h: '100%',
    fill: { color: COLORS.white, transparency: 30 },
    line: { type: 'none' }
  });

  // Title (centered, large, white)
  slid.addText(slide.title || 'Training Program', {
    x: 0.5, y: 2.0,
    w: 8.5, h: 1.2,
    fontSize: 48,
    bold: true,
    color: COLORS.white,
    fontFace: 'Calibri',
    align: 'center',
    valign: 'middle',
    wrap: true,
    lineSpacing: 32
  });

  // Subtitle (if exists)
  if (slide.subtitle) {
    slid.addText(slide.subtitle, {
      x: 0.5, y: 3.3,
      w: 8.5, h: 0.6,
      fontSize: 18,
      bold: false,
      color: COLORS.white,
      fontFace: 'Calibri',
      align: 'center',
      wrap: true
    });
  }

  // Label
  slid.addText('PROFESSIONAL TRAINING', {
    x: 0.5, y: 4.2,
    w: 8.5, h: 0.3,
    fontSize: 11,
    bold: true,
    color: 'rgba(255, 255, 255, 0.8)',
    fontFace: 'Calibri',
    align: 'center'
  });

  // Cover image (bottom-right, if exists)
  if (slide.generated_image_url) {
    slid.addImage({
      path: slide.generated_image_url,
      x: 5.0, y: 4.5,
      w: 4.5, h: 2.5
    });
  }

  // Footer date
  slid.addText(new Date().toLocaleDateString('en-GB'), {
    x: 0.5, y: 6.8,
    w: 8.5, h: 0.25,
    fontSize: 10,
    color: 'rgba(255, 255, 255, 0.6)',
    fontFace: 'Calibri',
    align: 'right'
  });
}

// ============================================================================
// MAIN EXPORT HANDLER
// ============================================================================
export default async function handler(req, res) {
  // CORS & method check
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { presentation } = req.body;

    if (!presentation || !presentation.slides) {
      return res.status(400).json({ error: 'Missing presentation or slides' });
    }

    const slides = presentation.slides;
    if (!Array.isArray(slides) || slides.length === 0) {
      return res.status(400).json({ error: 'Presentation must have at least 1 slide' });
    }

    // Initialize PptxGenJS
    const prs = new PptxGenJS();
    prs.defineLayout({ name: 'LAYOUT1', width: 10, height: 7.5 });
    prs.defineLayout({ name: 'BLANK', width: 10, height: 7.5 });

    const brandColor = presentation.brandKit?.primary || COLORS.copper;

    // Generate slides
    slides.forEach((slide, idx) => {
      const slideNum = idx + 1;
      const layoutType = slide.layout_type || 'standard_learning';

      if (layoutType === 'cover' || idx === 0) {
        addCoverSlide(prs, slide, brandColor);
      } else if (layoutType === 'safety_critical' || (slide.safety_callouts && slide.safety_callouts.length > 0)) {
        addSafetyCriticalSlide(prs, slide, slideNum, brandColor);
      } else if (layoutType === 'process_checklist' || (slide.quality_callouts && slide.quality_callouts.length > 0)) {
        addProcessChecklistSlide(prs, slide, slideNum, brandColor);
      } else {
        addStandardLearningSlide(prs, slide, slideNum, brandColor);
      }

      // Add presenter notes with avatar script + video reference (if available)
      const lastSlide = prs.slides[prs.slides.length - 1];
      if (slide.presenter_notes || slide.avatar_script || slide.avatar_video_url) {
        let notesText = '';
        if (slide.presenter_notes) notesText += slide.presenter_notes;
        if (slide.avatar_script) notesText += `\n\n[AVATAR NARRATION]\n${slide.avatar_script}`;
        if (slide.avatar_video_url) notesText += `\n\n[AVATAR VIDEO]\n${slide.avatar_video_url}`;
        if (notesText) lastSlide.notesText = notesText;
      }
    });

    // Generate and encode PPTX
    const buffer = await prs.write({ outputType: 'arraybuffer' });

    // Convert to base64
    const base64 = Buffer.from(buffer).toString('base64');

    const fileName = `${(presentation.title || 'Presentation').replace(/[^a-zA-Z0-9]/g, '-')}-${new Date().toISOString().split('T')[0]}.pptx`;

    return res.status(200).json({
      success: true,
      export_version: 'PHASE_2_BRANDED_HANDLER_2026_06_14',
      file_name: fileName,
      pptx_base64: base64,
      slide_count: slides.length,
      pptx_size_bytes: buffer.byteLength,
      generated_at: new Date().toISOString()
    });

  } catch (error) {
    console.error('[PPTX Export Error]', error);
    return res.status(500).json({
      error: error.message,
      details: error.stack
    });
  }
}
