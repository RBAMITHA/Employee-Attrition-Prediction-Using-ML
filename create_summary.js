const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, BorderStyle, WidthType, ShadingType, HeadingLevel,
        LevelFormat } = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const borders = { top: border, bottom: border, left: border, right: border };

const divider = new Paragraph({
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '2E4A8A', space: 1 } },
  spacing: { before: 120, after: 120 }
});

function heading(text, size=26, color='1F3A6E') {
  return new Paragraph({
    spacing: { before: 240, after: 100 },
    children: [new TextRun({ text, bold: true, size, color, font: 'Arial' })]
  });
}

function para(text, options={}) {
  return new Paragraph({
    spacing: { before: 60, after: 80 },
    children: [new TextRun({ text, size: 22, font: 'Arial', ...options })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: 'Arial' })]
  });
}

function statRow(label, value, highlight=false) {
  const fill = highlight ? 'FFF3CD' : 'FFFFFF';
  return new TableRow({
    children: [
      new TableCell({
        borders, width: { size: 5200, type: WidthType.DXA },
        shading: { fill, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: label, size: 21, font: 'Arial', bold: true })] })]
      }),
      new TableCell({
        borders, width: { size: 4160, type: WidthType.DXA },
        shading: { fill, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: value, size: 21, font: 'Arial' })] })]
      }),
    ]
  });
}

const doc = new Document({
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
    }]
  },
  styles: {
    default: { document: { run: { font: 'Arial', size: 22 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1260, bottom: 1080, left: 1260 }
      }
    },
    children: [
      // Header
      new Paragraph({
        alignment: AlignmentType.CENTER,
        shading: { fill: '1F3A6E', type: ShadingType.CLEAR },
        spacing: { before: 0, after: 0 },
        children: [
          new TextRun({ text: '', break: 1 }),
          new TextRun({ text: 'EMPLOYEE ATTRITION ANALYSIS', bold: true, size: 32, color: 'FFFFFF', font: 'Arial' }),
          new TextRun({ text: '', break: 1 }),
          new TextRun({ text: 'A Summary Report for HR Leadership', size: 22, color: 'BDD7EE', font: 'Arial' }),
          new TextRun({ text: '', break: 1 }),
          new TextRun({ text: 'Prepared by: Bamitha R  |  Date: June 2026', size: 20, color: 'BDD7EE', font: 'Arial', italics: true }),
          new TextRun({ text: '', break: 1 }),
        ]
      }),

      new Paragraph({ spacing: { before: 200, after: 0 }, children: [new TextRun('')] }),

      // What we did
      heading('What We Did'),
      divider,
      para('We analysed data from 1,470 employees — covering their salaries, job roles, department, work hours, satisfaction ratings, and whether they eventually left the company. Using this data, we built a computer model that can identify which employees are most likely to leave in the future, before they actually resign.'),
      para('Think of it like a "flight risk" score — the model looks at patterns in the data and flags employees who share characteristics with people who have left in the past.'),

      new Paragraph({ spacing: { before: 160, after: 0 }, children: [new TextRun('')] }),

      // Key findings
      heading('What We Found — Key Numbers'),
      divider,

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [5200, 4160],
        rows: [
          statRow('Total employees studied', '1,470'),
          statRow('Employees who left (attrition rate)', '16.2% — roughly 1 in 6', true),
          statRow('Highest-risk department', 'Sales (~21% exit rate)', true),
          statRow('Highest-risk job role', 'Sales Representative'),
          statRow('Employees at highest risk by tenure', 'Those in first 0–2 years', true),
          statRow('Model accuracy (ROC-AUC score)', '~0.83 out of 1.00 — strong predictive power'),
        ]
      }),

      new Paragraph({ spacing: { before: 160, after: 0 }, children: [new TextRun('')] }),

      heading('The 3 Biggest Reasons Employees Leave'),
      divider,

      para('After running the analysis, three factors stood out above all others as the strongest warning signs:'),

      new Paragraph({ spacing: { before: 80, after: 0 }, children: [new TextRun('')] }),

      para('1.  Low Monthly Salary', { bold: true, color: 'C0392B' }),
      para('   Employees earning significantly below the company average are far more likely to resign. This was the single strongest predictor in our model. Employees who left earned a median salary roughly 40-50% lower than those who stayed.'),

      para('2.  Working Overtime Regularly', { bold: true, color: 'C0392B' }),
      para('   Employees who are frequently required to work beyond normal hours show a much higher exit rate. This points to burnout. Importantly, even well-paid employees who worked heavy overtime still chose to leave — showing that money alone cannot compensate for exhaustion.'),

      para('3.  Being New to the Company (0–2 Years Tenure)', { bold: true, color: 'C0392B' }),
      para('   The first two years at a company are the most dangerous window for attrition. New joiners who do not feel welcomed, valued, or see a clear growth path tend to start looking elsewhere quickly. Attrition in this group is disproportionately high compared to employees who have been with the company for 5+ years.'),

      new Paragraph({ spacing: { before: 160, after: 0 }, children: [new TextRun('')] }),

      heading('Our Recommendations'),
      divider,

      para('Recommendation 1 — Targeted Salary Review for Sales & Junior Roles', { bold: true }),
      para('We recommend an immediate pay benchmarking exercise for Sales Representatives and entry-level staff earning in the bottom 25% of their pay band. A focused 10–15% salary correction for this group is likely far cheaper than the cost of recruiting, hiring, and training a replacement — which typically costs 50–200% of an employee\'s annual salary.'),

      new Paragraph({ spacing: { before: 80, after: 0 }, children: [new TextRun('')] }),

      para('Recommendation 2 — Overtime Alerts + a New Joiner Engagement Programme', { bold: true }),
      para('We recommend two parallel actions: First, HR should receive an automated monthly alert for any employee who has logged more than 20% overtime hours in a 30-day period — this triggers a voluntary wellness check-in. Second, introduce a structured milestone programme for all new joiners at 30, 90, and 365 days — including a dedicated buddy, a career development conversation, and a feedback session. This directly targets the high-risk 0–2 year window.'),

      new Paragraph({ spacing: { before: 160, after: 0 }, children: [new TextRun('')] }),

      heading('A Note of Caution'),
      divider,
      para('This model is a decision-support tool, not a final verdict. It highlights patterns and probabilities based on historical data. It cannot predict individual behaviour with certainty, and should never be used to penalise or make employment decisions about specific employees without a human conversation first.'),
      para('HR leadership should treat a high-risk flag as the start of a supportive conversation — not a disciplinary one.'),

      new Paragraph({ spacing: { before: 200, after: 0 }, children: [new TextRun('')] }),

      // Footer line
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: '1F3A6E', space: 1 } },
        spacing: { before: 120, after: 0 },
        children: [new TextRun({ text: 'Confidential — For Internal HR Use Only  |  InternPe Internship Project — Week 2  |  Bamitha R', size: 18, color: '888888', font: 'Arial', italics: true })]
      })
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/home/claude/EmployeeAttrition_Bamitha/summary.docx', buf);
  console.log('summary.docx created!');
}).catch(e => { console.error(e); process.exit(1); });
