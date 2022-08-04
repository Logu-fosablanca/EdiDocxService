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
  ImageRun,
  WidthType,
  Table,
  TableCell,
  TableRow,
  Header,
  BorderStyle,
  SectionType,
  TextWrappingType,
  TextWrappingSide,
  PageBreak,
} = docx;
import path from "path";
import { fileURLToPath } from "url";
import pkg from "file-saver";
const { saveAs } = pkg;

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
  console.log(params.segmentUsage);

  let segmentUsage = JSON.parse(JSON.stringify(params.segmentUsage));
  let numberOfElementsInSegment = JSON.parse(
    JSON.stringify(params.numberOfElementsInSegment)
  );
  let elementUsageDefs = JSON.parse(JSON.stringify(params.elementUsageDefs));
  let segmentText = JSON.parse(JSON.stringify(params.segmentText));
  let elementCode = JSON.parse(JSON.stringify(params.code));
  let y, z, a;
  let tempElementCode = elementCode;
  let testversion = version.replace(/\//g, "_");

  let fileName =  businessPartnerText + "_" + transactionSet + "_" + testversion + ".docx";

  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", "attachment; filename=" + fileName);

  let filePath = __dirname + "/EDIFiles/docx/" + fileName;
  let test = __dirname + "/EDIFiles/docx/" + "example.docx";
  console.log("end of file");
  console.log(req);

  let x;
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

  // This Displays the Segment Usage part of the Application

  let segmentData = Object.values(segmentUsage);
  console.log(segmentData);
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
  });

  // This the Table Header for the Segment Usage Table in the Word Document (docx)

  let SegmentHeader = [
    new TableRow({
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
    }),
  ];

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
    rows: [...SegmentHeader, ...segmentArray],
  });

  /* 

 For each segment part we will be displaying a certain set of the Elemts Used in the Application 
 So we will be creating a dictionary that stores the values and its segment id as key value pair
 
*/

  let segmentelementMap = {};
  for (x in segmentData) {
    segmentelementMap[segmentData[x].SegmentID] = segmentData[x].Description;
  }

  let elementData = Object.values(elementUsageDefs);
  let elementDataIndex = Object.values(elementData);
  console.log(elementDataIndex);
  let elementDataHead = [];
  for (let x in elementDataIndex) {
    for (y in elementDataIndex[x]) {
      elementDataHead.push(y);
    }
    break;
  }

  let SegmentIDlist = [];
  let tablelist = {};

  for (let x in elementDataIndex) {
    console.log(" New Dictionary " + x);
    for (let y in elementDataIndex[x]) {
      console.log("the object " + JSON.stringify(elementDataIndex[x][y]));
      if (SegmentIDlist.indexOf(elementDataIndex[x][y].SegmentID) == -1) {
        SegmentIDlist.push(elementDataIndex[x][y].SegmentID);
      }

      if (tablelist[elementDataIndex[x][y].SegmentID] == undefined) {
        tablelist[elementDataIndex[x][y].SegmentID] = [];
      }
      tablelist[elementDataIndex[x][y].SegmentID].push(elementDataIndex[x][y]);

      for (let z in elementDataIndex[x][y]) {
        console.log(elementDataIndex[x][y].SegmentID);
      }
    }
  }
  console.log(tablelist);

  /*
The section should have
  ...arrray of tables 


  and push the entire table into the table array 
  
    each table'function will return a table from the respective segmentid  
             1) with ElementTableHeader 
             2) Table Data   [ Store all these in an array and merge them up  ][...ElementTableHeader, ...TableData] with table data is equivalent to element = element.map{
              return new TableRow{
                  children: [....]
              }
             }


*/
  function ElementTableRowGenerator(segmentId, object) {
    let tablerows = object.map((element) => {
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
    });
    return tablerows;
  }

  let ElementTableHeader = [
    new TableRow({
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
    }),
  ];

  function ElementsTableGenerator(object) {
    let document = [];
    Object.keys(object).map(function (segmentId, ElemtsArray) {
      console.log(object[segmentId]);
      document.push(
        new Paragraph({
          text: segmentId + " " + segmentelementMap[segmentId],
          bold: true,
          run: {
            italics: true,
            color: "999999",
          },
        })
      );
      document.push(
        new Paragraph({
          text: "Element Summary",
          bold: true,
        })
      );
      document.push(
        new Table({
          rows: [].concat(
            ElementTableHeader,
            ElementTableRowGenerator(segmentId, object[segmentId])
          ),
        })
      );
      document.push(
        new Paragraph({
          children: [new PageBreak()],
        })
      );
    });
    return document;
  }

  let ElementDocumenttables = ElementsTableGenerator(tablelist);

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
                  // image,
                  new ImageRun({
                    data: fs.readFileSync("./assets/logo.jpg"),
                    transformation: {
                      width: 70,
                      height: 70,
                    },
                    floating: {
                      horizontalPosition: {
                        offset: 514400,
                      },
                      verticalPosition: {
                        offset: 514400,
                      },
                      wrap: {
                        type: TextWrappingType.SQUARE,
                        side: TextWrappingSide.BOTH_SIDES,
                      },
                      margins: {
                        top: 201440,
                        bottom: 201440,
                      },
                    },
                  }),
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
            text: "Segmentation - Elements Map",
            bold: true,
          }),
          SegmentDatatable,
        ],
      },
    ],
  });

  doc.addSection({
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
    children: [...ElementDocumenttables],
  });

  // Used to export the file into a .docx file

  Packer.toBuffer(doc).then((buffer) => {
    fs.createWriteStream(filePath);
    fs.writeFileSync(fileName, buffer);
  });

  res.send("ok");
}
