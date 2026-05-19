import PptxGenJS from 'pptxgenjs';

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { visits, reportType, reportPeriod, reportYear } = req.body;

    if (!visits || !reportType || !reportPeriod || !reportYear) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    const colors = { dark: '#0d1628', cyan: '#00c8ff', orange: '#f39c12', green: '#27ae60', white: '#ffffff', text: '#cccccc' };
    const font = 'Montserrat';

    const totalPromotars = visits.reduce((sum, v) => sum + (v.promotersCount || 0), 0);
    const totalAmbassadors = visits.reduce((sum, v) => sum + (v.ambassadorsCount || 0), 0);
    const totalStaff = visits.reduce((sum, v) => sum + (v.staffCount || 0), 0);
    const totalTrained = totalPromotars + totalAmbassadors + totalStaff;
    const totalSellout = visits.reduce((sum, v) => sum + (v.sellout || 0), 0);
    const totalTarget = visits.reduce((sum, v) => sum + (v.target || 0), 0);
    const achievement = totalTarget > 0 ? ((totalSellout / totalTarget) * 100).toFixed(1) : 0;

    let periodName = '';
    if (reportType === 'weekly') periodName = `Week ${reportPeriod}`;
    else if (reportType === 'monthly') {
      const months = ['', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      periodName = months[reportPeriod];
    } else if (reportType === 'quarterly') periodName = `Q${reportPeriod}`;

    const pres = new PptxGenJS();
    pres.defineLayout({ name: 'BLANK', width: 10, height: 7.5 });

    // SLIDE 1: COVER
    let slide = pres.addSlide();
    slide.background = { color: colors.dark };
    slide.addText('ZTE TRAINING REPORT', { x: 0.5, y: 2, w: 9, h: 0.8, fontSize: 48, bold: true, color: colors.cyan, fontFace: font, align: 'center' });
    slide.addText(periodName + ' · ' + reportYear, { x: 0.5, y: 3, w: 9, h: 0.5, fontSize: 32, color: colors.white, fontFace: font, align: 'center' });
    slide.addText('Trainer: Simone Stasiano', { x: 0.5, y: 4, w: 9, h: 0.4, fontSize: 16, color: colors.text, fontFace: font, align: 'center' });

    // SLIDE 2: OVERVIEW
    slide = pres.addSlide();
    slide.background = { color: colors.dark };
    slide.addText('OVERVIEW', { x: 0.5, y: 0.4, w: 9, h: 0.5, fontSize: 28, bold: true, color: colors.cyan, fontFace: font });

    const kpiData = [
      { label: 'Total Visits', value: visits.length, color: colors.cyan },
      { label: 'Total Trained', value: totalTrained, color: colors.cyan },
      { label: 'Promoters', value: totalPromotars, color: colors.orange },
      { label: 'Ambassadors', value: totalAmbassadors, color: colors.green },
      { label: 'Store Staff', value: totalStaff, color: colors.white },
      { label: 'Total Sell-out', value: totalSellout, color: colors.cyan },
      { label: 'Total Target', value: totalTarget, color: colors.orange },
      { label: 'Achievement %', value: achievement + '%', color: achievement >= 80 ? colors.green : achievement >= 60 ? colors.orange : '#e74c3c' }
    ];

    let kpiX = 0.5, kpiY = 1.2;
    for (let i = 0; i < kpiData.length; i++) {
      const kpi = kpiData[i];
      const boxW = 2, boxH = 1;
      if (kpiX + boxW > 10) { kpiX = 0.5; kpiY += boxH + 0.3; }

      slide.addShape(pres.ShapeType.rect, { x: kpiX, y: kpiY, w: boxW, h: boxH, fill: { color: colors.dark }, line: { color: kpi.color, width: 2 } });
      slide.addText(kpi.label, { x: kpiX + 0.1, y: kpiY + 0.1, w: boxW - 0.2, h: 0.35, fontSize: 11, color: colors.text, fontFace: font, bold: true });
      slide.addText(String(kpi.value), { x: kpiX + 0.1, y: kpiY + 0.5, w: boxW - 0.2, h: 0.35, fontSize: 18, color: kpi.color, fontFace: font, bold: true, align: 'center' });

      kpiX += boxW + 0.3;
    }

    // SLIDE 3+: PER OGNI VISITA
    visits.forEach((visit, idx) => {
      slide = pres.addSlide();
      slide.background = { color: colors.dark };

      slide.addText(`VISIT ${idx + 1} - ${visit.storeName}`, { x: 0.5, y: 0.4, w: 9, h: 0.5, fontSize: 24, bold: true, color: colors.cyan, fontFace: font });
      slide.addText(visit.visitDate, { x: 0.5, y: 1, w: 9, h: 0.3, fontSize: 12, color: colors.text, fontFace: font });

      slide.addText('DETAILS', { x: 0.5, y: 1.5, w: 4.2, h: 0.35, fontSize: 12, bold: true, color: colors.cyan, fontFace: font });
      let detailY = 2;
      const details = [
        { label: 'Location:', value: visit.storeLocation },
        { label: 'Promoters:', value: visit.promotersCount },
        { label: 'Ambassadors:', value: visit.ambassadorsCount },
        { label: 'Staff:', value: visit.staffCount },
        { label: 'Reason:', value: (visit.reasonForVisit || []).join(', ') || 'N/A' },
        { label: 'Delivery:', value: (visit.deliveryMode || []).join(', ') || 'N/A' }
      ];
      details.forEach(d => {
        slide.addText(`${d.label}`, { x: 0.5, y: detailY, w: 1.5, h: 0.25, fontSize: 10, bold: true, color: colors.orange, fontFace: font });
        slide.addText(String(d.value), { x: 2.1, y: detailY, w: 2.6, h: 0.25, fontSize: 10, color: colors.white, fontFace: font });
        detailY += 0.3;
      });

      slide.addText('PERFORMANCE', { x: 5.3, y: 1.5, w: 4.2, h: 0.35, fontSize: 12, bold: true, color: colors.cyan, fontFace: font });
      let perfY = 2;
      const perfData = [
        { label: 'Sell-out:', value: visit.sellout, unit: 'units' },
        { label: 'Target:', value: visit.target, unit: 'units' },
        { label: 'Achievement:', value: visit.target > 0 ? ((visit.sellout / visit.target) * 100).toFixed(1) : '—', unit: '%' }
      ];
      perfData.forEach(p => {
        slide.addText(`${p.label}`, { x: 5.3, y: perfY, w: 1.5, h: 0.25, fontSize: 10, bold: true, color: colors.orange, fontFace: font });
        slide.addText(`${p.value} ${p.unit}`, { x: 6.9, y: perfY, w: 2.6, h: 0.25, fontSize: 10, color: colors.white, fontFace: font });
        perfY += 0.3;
      });

      if (visit.notes) {
        slide.addText('NOTES', { x: 0.5, y: 5, w: 9, h: 0.3, fontSize: 11, bold: true, color: colors.cyan, fontFace: font });
        slide.addText(visit.notes, { x: 0.5, y: 5.35, w: 9, h: 1.5, fontSize: 9, color: colors.text, fontFace: font, wrap: true });
      }

      if (visit.visitPhotos && visit.visitPhotos.length > 0) {
        slide = pres.addSlide();
        slide.background = { color: colors.dark };
        slide.addText(`VISIT ${idx + 1} - PHOTO EVIDENCE`, { x: 0.5, y: 0.4, w: 9, h: 0.5, fontSize: 20, bold: true, color: colors.cyan, fontFace: font });

        const photosPerRow = visit.visitPhotos.length <= 2 ? visit.visitPhotos.length : 3;
        const photoW = (9 - 0.5 * (photosPerRow + 1)) / photosPerRow;
        const photoH = photoW * 0.75;
        let photoX = 0.5, photoY = 1.2;

        visit.visitPhotos.forEach((photo, pIdx) => {
          if (pIdx > 0 && pIdx % photosPerRow === 0) { photoX = 0.5; photoY += photoH + 0.3; }
          try {
            slide.addImage({ data: photo, x: photoX, y: photoY, w: photoW, h: photoH });
          } catch (e) {
            slide.addShape(pres.ShapeType.rect, { x: photoX, y: photoY, w: photoW, h: photoH, fill: { color: '#333' } });
          }
          photoX += photoW + 0.3;
        });
      }
    });

    // Convert to buffer, then to base64
    const buffer = await pres.write({ outputType: 'arraybuffer' });
    const base64 = Buffer.from(buffer).toString('base64');
    const filename = `ZTE_Report_${reportType}_${reportPeriod}_${reportYear}.pptx`;

    // Return base64 as JSON
    res.status(200).json({ pptx: base64, filename: filename });

  } catch (error) {
    console.error('PPTX generation error:', error);
    res.status(500).json({ error: 'Failed to generate PPTX: ' + error.message });
  }
}
