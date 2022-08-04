import docx from "docx";
const { Table, Packer, TableRow, TableCell, Paragraph,Document } = docx;
import * as fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
// import JSON from "JSON";

const data = [
  {
    1: {
      _id: "629f3e7387b28e37a567ab7c",
      Agency: "A",
      Version: "ADJL01",
      SegmentID: "CID",
      ElementID: "C001",
      Release: "1",
      Position: "1",
      RequirementDesignator: "M",
      GroupRequirementDesignatorID: "SCID",
      SubElementReqDesignator: "M",
      GroupReqDesignator: "M",
      GroupBeginEnd: "G",
      Type: "ID",
      Description: "COMPANY IDENTIFIER",
      MinimumLength: "6",
      MaximumLength: "6",
      CompositeElement: "FALSE",
      RepeatFactor: "1",
      SegmentPosition: "1",
    },
    2: {
      _id: "629f3e7387b28e37a567ab7d",
      Agency: "A",
      Version: "ADJL01",
      SegmentID: "CID",
      ElementID: "SCI2",
      Release: "1",
      Position: "2",
      RequirementDesignator: "M",
      GroupRequirementDesignatorID: "SCID",
      SubElementReqDesignator: "M",
      GroupReqDesignator: "M",
      Type: "AN",
      Description: "SENDING COMPANY/OPERATOR ID",
      MinimumLength: "1",
      MaximumLength: "10",
      CompositeElement: "FALSE",
      RepeatFactor: "0",
      SegmentPosition: "1",
    },
    3: {
      _id: "629f3e7387b28e37a567ab7e",
      Agency: "A",
      Version: "ADJL01",
      SegmentID: "CID",
      ElementID: "C001",
      Release: "1",
      Position: "3",
      RequirementDesignator: "M",
      GroupRequirementDesignatorID: "RCID",
      SubElementReqDesignator: "M",
      GroupReqDesignator: "M",
      GroupBeginEnd: "G",
      Type: "ID",
      Description: "COMPANY IDENTIFIER",
      MinimumLength: "6",
      MaximumLength: "6",
      CompositeElement: "FALSE",
      RepeatFactor: "1",
      SegmentPosition: "1",
    },
    4: {
      _id: "629f3e7387b28e37a567ab7f",
      Agency: "A",
      Version: "ADJL01",
      SegmentID: "CID",
      ElementID: "RCI2",
      Release: "1",
      Position: "4",
      RequirementDesignator: "M",
      GroupRequirementDesignatorID: "RCID",
      SubElementReqDesignator: "M",
      GroupReqDesignator: "M",
      Type: "AN",
      Description: "RECEIVING COMPANY/OPERATOR ID",
      MinimumLength: "1",
      MaximumLength: "10",
      CompositeElement: "FALSE",
      RepeatFactor: "0",
      SegmentPosition: "1",
    },
  },
  {
    1: {
      _id: "629f3e7387b28e37a567ac6d",
      Agency: "A",
      Version: "ADJL01",
      SegmentID: "PLH",
      ElementID: "PLNO",
      Release: "1",
      Position: "1",
      RequirementDesignator: "M",
      SubElementReqDesignator: "M",
      Type: "AN",
      Description: "POLICY NUMBER",
      MinimumLength: "1",
      MaximumLength: "20",
      CompositeElement: "FALSE",
      RepeatFactor: "1",
      SegmentPosition: "2",
    },
    2: {
      _id: "629f3e7387b28e37a567ac75",
      Agency: "A",
      Version: "ADJL01",
      SegmentID: "PLH",
      ElementID: "OVS2",
      Release: "1",
      Position: "9",
      RequirementDesignator: "M",
      GroupRequirementDesignatorID: "OVSN",
      SubElementReqDesignator: "M",
      GroupReqDesignator: "M",
      Type: "N0",
      Description: "NUMBER",
      MinimumLength: "2",
      MaximumLength: "2",
      CompositeElement: "FALSE",
      RepeatFactor: "0",
      SegmentPosition: "2",
    },
  },
];

let ElementTableHeader=  [new TableRow({
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

let SegmentIDlist = [];
let tablelist = {};

for (let x in data) {
  console.log(" New Dictionary " + x);
  for (let y in data[x]) {
    console.log("the object " + JSON.stringify(data[x][y]));
    if (SegmentIDlist.indexOf(data[x][y].SegmentID) == -1) {
      SegmentIDlist.push(data[x][y].SegmentID);
    }

    if (tablelist[data[x][y].SegmentID] == undefined) {
      tablelist[data[x][y].SegmentID] = [];
    }
    tablelist[data[x][y].SegmentID].push(data[x][y]);

    for (let z in data[x][y]) {
      console.log(data[x][y].SegmentID);
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
function ElementTableRowGenerator ( segmentId,object) {
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
  })
  return tablerows;
}



function ElementsTableGenerator ( object ){

let document =[];
 Object.keys(object).map(function(segmentId,ElemtsArray) {
    console.log(object[segmentId]);
    document.push(new Paragraph({
      text: segmentId ,
      bold: true,
    }))
    document.push(new Table({
      rows: [].concat(ElementTableHeader, ElementTableRowGenerator(segmentId, object[segmentId])),
    }))
   
  })
  return document;
}




let ElementDocumenttables = ElementsTableGenerator(tablelist);


let doc = new Document({
  sections: [
{    children: [...ElementDocumenttables],}
  ]
});



Packer.toBuffer(doc).then((buffer) => {
  // fs.createWriteStream(filePath);
  fs.writeFileSync('xyz.docx', buffer);
});


console.log((ElementDocumenttables));




