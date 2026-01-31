const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        HeadingLevel, BorderStyle, WidthType, ShadingType, AlignmentType,
        LevelFormat } = require('docx');
const fs = require('fs');

// Helper function for creating paragraphs with spacing
function spacedPara(text, options = {}) {
  return new Paragraph({
    spacing: { after: 200 },
    ...options,
    children: [new TextRun(text)]
  });
}

function boldPara(text, options = {}) {
  return new Paragraph({
    spacing: { after: 200 },
    ...options,
    children: [new TextRun({ text, bold: true })]
  });
}

// Table border style
const border = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

// Header cell styling
function headerCell(text, width) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: "D9E2F3", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, size: 22 })] })]
  });
}

function dataCell(content, width) {
  const children = typeof content === 'string' 
    ? [new Paragraph({ children: [new TextRun({ text: content, size: 22 })] })]
    : content;
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children
  });
}

// Schedule data
const schedule = [
  { week: "1", tuesday: "Course Introduction: Data, Evidence, and Argument in the Humanities", thursday: "Workshop: Setting Up Your Environment; Baseline Reflection assigned" },
  { week: "2", tuesday: "Module 1: Textual Data—What Does It Mean to Count Words?", thursday: "Lab 1a: Building and Exploring a Text Corpus" },
  { week: "3", tuesday: "Frequency, Distribution, and the Limits of Counting", thursday: "Lab 1b: Visualizing Textual Patterns; Module 1 Deliverable due" },
  { week: "4", tuesday: "Module 2: Tabular Data—Categories, Records, and What Gets Left Out", thursday: "Lab 2a: Cleaning and Structuring Historical Data" },
  { week: "5", tuesday: "Comparison Across Time and Groups", thursday: "Lab 2b: Comparative Visualizations; Module 2 Deliverable due" },
  { week: "6", tuesday: "Module 3: Spatial Data—Place as Evidence", thursday: "Lab 3a: From Coordinates to Questions" },
  { week: "7", tuesday: "Mapping Arguments, Not Just Locations", thursday: "Lab 3b: Geographic Visualization; Module 3 Deliverable due" },
  { week: "8", tuesday: "BREAK", thursday: "BREAK" },
  { week: "9", tuesday: "Module 4: Relational Data—Networks as Interpretation", thursday: "Lab 4a: Building and Querying Networks" },
  { week: "10", tuesday: "From Pattern to Argument: Synthesizing Methods", thursday: "Lab 4b: Interactive Visualization; Module 4 Deliverable due; Midterm Reflection due" },
  { week: "11", tuesday: "Independent Project: Proposals and Data Assessment", thursday: "Workshop: Peer Feedback on Proposals" },
  { week: "12", tuesday: "Workshop: Data Acquisition and Preparation", thursday: "Workshop: Troubleshooting and Consultation" },
  { week: "13", tuesday: "Workshop: Exploratory Analysis", thursday: "Workshop: Identifying Patterns and Problems" },
  { week: "14", tuesday: "Workshop: Building Your Visual Argument", thursday: "Workshop: Iteration and Revision" },
  { week: "15", tuesday: "Workshop: Documentation and Reflection", thursday: "Workshop: Dress Rehearsal Presentations" },
  { week: "16", tuesday: "Final Presentations; Final Reflection due", thursday: "" }
];

// Assignments data
const assignments = [
  { name: "Baseline Reflection", desc: "Students articulate their incoming skills, assumptions, and goals for the course." },
  { name: "Lab 1a: Building and Exploring a Text Corpus", desc: "Hands-on work acquiring, cleaning, and performing initial exploration of textual data." },
  { name: "Lab 1b: Visualizing Textual Patterns", desc: "Creating frequency-based visualizations and interpreting what they reveal and obscure." },
  { name: "Module 1 Deliverable: Textual Analysis", desc: "A visualization of textual data accompanied by a brief interpretive statement addressing choices made and limitations encountered." },
  { name: "Lab 2a: Cleaning and Structuring Historical Data", desc: "Working with tabular historical records; addressing missing data, inconsistent categories, and documentation." },
  { name: "Lab 2b: Comparative Visualizations", desc: "Building visualizations that compare across time periods or groups; small multiples, grouped charts, time series." },
  { name: "Module 2 Deliverable: Comparative Analysis", desc: "A comparative visualization with written reflection on categorization decisions and their interpretive consequences." },
  { name: "Lab 3a: From Coordinates to Questions", desc: "Geocoding, joining spatial data, and formulating place-based research questions." },
  { name: "Lab 3b: Geographic Visualization", desc: "Creating maps that argue rather than merely display; choropleth, point density, and layered approaches." },
  { name: "Module 3 Deliverable: Spatial Argument", desc: "A map-based visualization with accompanying text explaining how spatial representation shapes interpretation." },
  { name: "Lab 4a: Building and Querying Networks", desc: "Constructing network data from humanities sources; nodes, edges, and attributes." },
  { name: "Lab 4b: Interactive Visualization", desc: "Building visualizations that invite exploration; interactivity as a form of argument." },
  { name: "Module 4 Deliverable: Relational Analysis", desc: "A network or interactive visualization with reflection on what relational structure reveals." },
  { name: "Midterm Reflection with Proposed Grade", desc: "Students assess their learning so far, citing specific evidence from their work, and propose a grade with justification." },
  { name: "Independent Project Proposal", desc: "A brief document identifying the research question, data source(s), intended audience, and preliminary approach." },
  { name: "Final Project: Functional Component", desc: "A polished visualization or interactive piece designed for a specific audience, suitable for public presentation or portfolio." },
  { name: "Final Project: Reflective Documentation", desc: "A methodological narrative documenting choices, dead ends, revisions, and what the data can and cannot support." },
  { name: "Final Reflection with Proposed Grade", desc: "Students make an evidence-based argument for their course grade, drawing on all work produced and growth demonstrated." }
];

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 24 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 360, after: 240 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 300, after: 180 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 2 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: [
      // Title
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER,
        children: [new TextRun("Data Analysis and Visualization in the Humanities")]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        children: [new TextRun({ text: "Course Planning Document", italics: true, size: 24 })]
      }),

      // Course Rationale
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Course Rationale")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun({ text: "For curriculum committee review", italics: true, size: 22 })]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("This course addresses a critical gap in the Digital Culture and Data Analytics curriculum: the need for a methods course that bridges foundational technical skills and advanced application. Students entering the course will have completed introductory coursework in both statistics and programming (Python or R), giving them the technical vocabulary to work with data. What they lack is sustained practice in the interpretive and rhetorical dimensions of data work—the capacity to move from pattern recognition to meaningful argument.")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("The course is grounded in a conviction that distinguishes humanistic data analysis from its counterparts in business or the social sciences: that data should be used to test beliefs rather than defend them. This epistemological commitment shapes everything from the questions students learn to ask (\"What would change my mind?\") to the visualizations they learn to build (displays that show reasoning, not just conclusions). In an era when data visualizations circulate widely and often function as instruments of persuasion rather than inquiry, this critical orientation is both intellectually necessary and professionally valuable.")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("The course structure moves students from tightly scaffolded modules—where they work with shared datasets and develop common analytical vocabulary—toward increasing independence, culminating in a self-directed final project. This arc prepares students for the capstone course while also producing portfolio-ready work. The four modules expose students to different data types (textual, tabular/historical, spatial, and relational) and different visualization approaches, ensuring breadth while the final project allows depth.")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("Assessment follows an ungrading model in which students produce three reflections across the semester—baseline, midterm, and final—each requiring them to make evidence-based arguments about their own learning. This approach mirrors the course's intellectual commitments: just as students learn to hold their analytical claims accountable to evidence, they learn to hold their self-assessments accountable to the work they have produced.")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("The course is designed to be taught by rotating instructors with varied disciplinary backgrounds. Learning outcomes remain consistent, while specific datasets, examples, and disciplinary emphases can adapt to instructor expertise and student interests.")]
      }),

      // Catalog Description
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Catalog Description")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun({ text: "For student-facing materials", italics: true, size: 22 })]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun("Students will develop the skills to transform humanities research questions into data-driven arguments and compelling visualizations. Building on prior coursework in coding and statistics, this course emphasizes interpretation over implementation: learning to ask better questions, make defensible analytical choices, and communicate findings to diverse audiences. Through four modules covering textual, tabular, spatial, and relational data, students will practice the full cycle of humanistic data work—from formulating questions and preparing data to building visualizations that show reasoning rather than merely display conclusions. The course culminates in an independent project that produces both a functional, portfolio-ready visualization and reflective documentation of the methodological choices behind it. Prerequisite: WRIT 20833 or COSC 10603 or GEOG 30323; and MATH 10043 or INSC 20153.")]
      }),

      // Learning Outcomes
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Learning Outcomes")]
      }),
      new Paragraph({
        spacing: { after: 100 },
        children: [new TextRun("Upon successful completion of this course, students will be able to:")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 100 },
        children: [new TextRun("Formulate research questions that data can meaningfully address, identifying what evidence would be required to test or revise a belief.")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 100 },
        children: [new TextRun("Make and justify decisions about what to count, categorize, or exclude, articulating the interpretive consequences of these choices.")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 100 },
        children: [new TextRun("Clean and transform data for analysis, documenting preparation decisions and their rationale.")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 100 },
        children: [new TextRun("Choose appropriate visualization types for different analytical purposes, matching form to argument.")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 100 },
        children: [new TextRun("Construct visual arguments for specific audiences, designing displays that invite inquiry rather than foreclose it.")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 100 },
        children: [new TextRun("Critique data presentations—their own and others'—for rhetorical and evidentiary integrity, recognizing when visualizations obscure uncertainty or defend conclusions rather than test them.")]
      }),
      new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { after: 200 },
        children: [new TextRun("Document and communicate methodological choices, producing reflective accounts that make analytical reasoning visible.")]
      }),

      // Assignments and Projects
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun("Assignments and Projects")]
      }),
      
      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Reflections")]
      }),
      ...assignments.filter(a => a.name.includes("Reflection")).map(a => 
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 100 },
          children: [
            new TextRun({ text: a.name + ": ", bold: true }),
            new TextRun(a.desc)
          ]
        })
      ),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Module 1: Textual Data and Frequency Visualization")]
      }),
      ...assignments.filter(a => a.name.includes("1a") || a.name.includes("1b") || a.name.includes("Module 1 Deliverable")).map(a => 
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 100 },
          children: [
            new TextRun({ text: a.name + ": ", bold: true }),
            new TextRun(a.desc)
          ]
        })
      ),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Module 2: Tabular Data and Comparative Visualization")]
      }),
      ...assignments.filter(a => a.name.includes("2a") || a.name.includes("2b") || a.name.includes("Module 2 Deliverable")).map(a => 
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 100 },
          children: [
            new TextRun({ text: a.name + ": ", bold: true }),
            new TextRun(a.desc)
          ]
        })
      ),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Module 3: Spatial Data and Geographic Visualization")]
      }),
      ...assignments.filter(a => a.name.includes("3a") || a.name.includes("3b") || a.name.includes("Module 3 Deliverable")).map(a => 
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 100 },
          children: [
            new TextRun({ text: a.name + ": ", bold: true }),
            new TextRun(a.desc)
          ]
        })
      ),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Module 4: Relational Data and Network Visualization")]
      }),
      ...assignments.filter(a => a.name.includes("4a") || a.name.includes("4b") || a.name.includes("Module 4 Deliverable")).map(a => 
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 100 },
          children: [
            new TextRun({ text: a.name + ": ", bold: true }),
            new TextRun(a.desc)
          ]
        })
      ),

      new Paragraph({
        heading: HeadingLevel.HEADING_3,
        children: [new TextRun("Independent Project")]
      }),
      ...assignments.filter(a => a.name.includes("Proposal") || a.name.includes("Final Project")).map(a => 
        new Paragraph({
          numbering: { reference: "bullets", level: 0 },
          spacing: { after: 100 },
          children: [
            new TextRun({ text: a.name + ": ", bold: true }),
            new TextRun(a.desc)
          ]
        })
      ),

      // Course Schedule
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        pageBreakBefore: true,
        children: [new TextRun("Course Schedule")]
      }),
      new Paragraph({
        spacing: { after: 200 },
        children: [new TextRun({ text: "80-minute sessions, Tuesday/Thursday", italics: true, size: 22 })]
      }),

      // Schedule table
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [900, 4230, 4230],
        rows: [
          new TableRow({
            children: [
              headerCell("Week", 900),
              headerCell("Tuesday", 4230),
              headerCell("Thursday", 4230)
            ]
          }),
          ...schedule.map(row => 
            new TableRow({
              children: [
                dataCell(row.week, 900),
                dataCell(row.tuesday, 4230),
                dataCell(row.thursday, 4230)
              ]
            })
          )
        ]
      }),

      // Notes
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 400 },
        children: [new TextRun("Notes for Instructors")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { after: 100 },
        children: [new TextRun({ text: "Datasets and examples are flexible. ", bold: true }), new TextRun("The modules are organized by data type, but the specific corpora, archives, or sources used can and should reflect instructor expertise and student interests.")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { after: 100 },
        children: [new TextRun({ text: "Readings are purposeful and brief. ", bold: true }), new TextRun("Each reading should directly serve the work students are doing that week. Methodological how-tos, critical provocations, exemplary projects, and foundational texts on evidence and display all have a place.")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { after: 100 },
        children: [new TextRun({ text: "Workshop time increases across the semester. ", bold: true }), new TextRun("Early sessions include more structured discussion; later sessions shift toward intensive lab and workshop formats as students develop independence.")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { after: 100 },
        children: [new TextRun({ text: "The epistemological thread is explicit. ", bold: true }), new TextRun("Recurring attention to how we know what we claim to know—and what we owe others when we make claims—distinguishes this course from purely technical methods training.")]
      }),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { after: 100 },
        children: [new TextRun({ text: "Ungrading mirrors course content. ", bold: true }), new TextRun("The reflection structure asks students to make evidence-based arguments about their own learning, enacting the same epistemological commitments they are learning to apply to data.")]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/home/claude/data_viz_humanities_course.docx", buffer);
  console.log("Document created successfully.");
});
