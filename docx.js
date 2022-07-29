// import { text } from "body-parser";
import docx from "docx";
import * as filesaver from "file-saver";
import fs from "fs";
const {
  AlignmentType,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  TextRun,
  UnderlineType,
  WidthType,
  Table,
  TableCell,
  TableRow,
  Header,
  BorderStyle,
  SectionType,
} = docx;
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);

// 👇️ "/home/john/Desktop/javascript"
const __dirname = path.dirname(__filename);

export async function getDocx(req, res) {
  console.log(req.body);
  let params = req.body;
  let transactionSet = params.transactionSet;
  let version = params.version;
  console.log(version);
  let transactionDescription = params.transactionDescription;
  let transactionFunctionalGroup = params.transactionFunctionalGroup;
  let headingText = params.headingText;
  let footerText = params.footerText;
  footerText = footerText.split("$");
  let businessPartnerText = params.businessPartnerText;
  let numberOfHeadingSegments = params.numberOfHeadingSegments;

  let numberOfDetailSegments = params.numberOfDetailSegments;
  let numberOfSummarySegments = params.numberOfSummarySegments;
  let presentLoop = "";

  // params.segmentUsage=params.segmentUsage.substring(1,params.segmentUsage.length-1);
  console.log(params.segmentUsage);

  let segmentUsage = JSON.parse(params.segmentUsage);
  let numberOfElementsInSegment = JSON.parse(params.numberOfElementsInSegment);
  let elementUsageDefs = JSON.parse(params.elementUsageDefs);
  let segmentText = JSON.parse(params.segmentText);
  let elementCode = JSON.parse(params.code);
  let y, z, a;
  let tempElementCode = elementCode;
  let testversion = version.replace(/\//g, "_");

  let fileName =
    businessPartnerText + "_" + transactionSet + "_" + testversion + ".docx";

  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", "attachment; filename=" + fileName);

  let filePath = __dirname + "/EDIFiles/docx/" + fileName;
  let test = __dirname + "/EDIFiles/docx/" + "example.docx";
  console.log("end of file");
  console.log(req);

  for (x in tempElementCode) {
    for (y in tempElementCode[x]) {
      for (z in tempElementCode[x][y]) {
        console.log(tempElementCode[x][y][z]);
      }
    }
  }

  for (y in segmentText) {
    segmentText[y] = segmentText[y].split("$");
  }

  // Segment Table
  let segmentData = Object.values(segmentUsage);
  let segmentDataHead = [];

  let segmentArray = segmentData.map((segment) => {
    return new TableRow({
      children: [
        new TableCell({
          children: [
            new Paragraph({
              text: segment.Position,
              bold: true,
            }),
          ],
        }),
        new TableCell({
          children: [
            new Paragraph({
              text: segment.SegmentID,
              bold: true,
            }),
          ],
        }),
        new TableCell({
          children: [
            new Paragraph({
              text: segment.Description,
              bold: true,
            }),
          ],
        }),
        new TableCell({
          children: [
            new Paragraph({
              text: segment.RequirementDesignator,
              bold: true,
            }),
          ],
        }),
        new TableCell({
          children: [
            new Paragraph({
              text: segment.MaximumUsage,
              bold: true,
            }),
          ],
        }),
        new TableCell({
          children: [
            new Paragraph({
              text: segment.MaximumLoopRepeat,
              bold: true,
            }),
          ],
        }),
      ],
    });
  })


  let SegmentHeader= [new TableRow({
    children: [
      new TableCell({
        children: [
          new Paragraph({
            text: "Position",
            bold: true,
          }),
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: "SegmentId",
            bold: true,
          }),
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: "Segment Name",
            bold: true,
          }),
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: "Req NO",
            bold: true,
          }),
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: "MAX USE",
            bold: true,
          }),
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: "REPEAT",
            bold: true,
          }),
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: " ",
            bold: true,
          }),
        ],
      }),
    ],
})]

  for (let x in segmentData) {
    console.log(segmentData[x]);
    for (let y in segmentData[x]) {
      segmentDataHead.push(y);
    }
    break;
  }

  const SegmentDatatable = new Table({
    tableHeader: true,

    indent: {
      size: 600,
      type: WidthType.DXA,
    },
    borders: {
      left: {
        style: BorderStyle.DOT_DOT_DASH,
        size: 3,
        color: "00FF00",
      },
      right: {
        style: BorderStyle.DOT_DOT_DASH,
        size: 3,
        color: "ff8000",
      },
    },
    rows: [...SegmentHeader,...segmentArray],    
    
  });

  // Element Table
  let elementData = Object.values(elementUsageDefs);
  let elementDataIndex = Object.values(elementData[0]);
  console.log(elementDataIndex);
  let elementDataHead = [];

  let elementusafeheader= [new TableRow({
    children: [
      new TableCell({
        children: [
          new Paragraph({
            text: "Position",
            bold: true,
          }),
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: "SegmentID",
            bold: true,
          }), 
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: "Segment Name",
            bold: true,
          }),
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: "REQ NO",
            bold: true,
          }),
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: "Type",
            bold: true,
          }),
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: "Min/Max",
            bold: true,
          }),
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: "Notes",
            bold: true,
          }),
        ],
      }),
    ],
})]

//Sidebox Summaary 
let sideboxsummary = [new TableRow({
  children: [
    new TableCell({
      children: [
        new Paragraph({
        text: "POS",
        bold: true,})
      ],
    })
  ],
})]

let elementArray= elementDataIndex.map((element) => {
  return new TableRow({
    children: [

      new TableCell({
        children: [
          new Paragraph({
            text: element.SegmentID,
            bold: true,
          }),
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: element.ElementID,
            bold: true,
          }),
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: element.Description,
            bold: true,
          }),
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: element.RequirementDesignator,
            bold: true,
          }),
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: element.Type,
            bold: true,
          }),
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: element.MinimumLength + "/" + element.MaximumLength,
            bold: true,
          }),
        ],
      }),
      new TableCell({
        children: [
          new Paragraph({
            text: " ",
            bold: true,
          }),
        ],
      }),
    ],
  });
})


  for (let x in elementDataIndex) {
    for (y in elementDataIndex[x]) {
      elementDataHead.push(y);
    }
    break;
  }
  

  const ElementDatatable = new Table({

    rows: [...elementusafeheader,...elementArray],
  });

  const doc = new Document({
    creator: "EDI Document Generator",
    description: "EDI Document Generator  ",

    sections: [
      {
        properties: {
          type: SectionType.EVEN_PAGE,
        },
        // Heading part of the document

        headers: {
          default: new Header({
            children: [
              new Paragraph({
                spacing: {
                  before: 200,
                  after: 200,
                },
                shading: {
                  // type: ShadingType.REVERSE_DIAGONAL_STRIPE,
                  color: "0000FF",
                  fill: "#0000FF",
                },
                border: {
                  top: {
                    color: "#000000",
                    space: 1,
                    style: "single",
                    size: 6,
                  },
                  bottom: { 
                    color: "#000000",
                    space: 1,
                    style: "single",
                    size: 6,
                  },
                },
                children: [
                  new TextRun(transactionSet),
                  new Paragraph(
                    "VER." + version + " " + transactionDescription
                  ),
                  new Paragraph("Business Partner: " + businessPartnerText),
                ],
              }),
            ],
          }),
        },

        // Segmentation Part

        children: [
          new Paragraph({
            text: "Segmentation",
            bold: true,
          }),
          SegmentDatatable],
      },
    ],
  });
  
 
  doc.addSection(
    {
      // Element Part
      shading: {
        color: "00FFFF",
        fill: "FF0000",
      },
      border: {
        top: {
          color: "auto",
          space: 1,
          style: "single",
          size: 6,
        },
        bottom: {
          color: "auto",
          space: 1,
          style: "single",
          size: 6,
        },
      },
      children: [
        new Paragraph({
          text: "Element Summary",
          bold: true,
        }),
        ElementDatatable,
      ],
    }
  ); 

  console.log(segmentArray.length);
  console.log(SegmentHeader.length);
  // Used to export the file into a .docx file

  Packer.toBuffer(doc).then((buffer) => {
    fs.createWriteStream(filePath);
    fs.writeFileSync(fileName, buffer);
  });
  // filesaver.saveAs
  res.send("ok");
}
