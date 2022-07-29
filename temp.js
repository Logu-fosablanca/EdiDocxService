import * as fs from "fs";
import docx from "docx"
const { Document, HeadingLevel,WidthType,TextRun,Packer, Paragraph, Table, TableCell, TableRow, VerticalAlign, TextDirection } = docx;

const data= {
    '1': {
      _id: '629f406787b28e37a57d2133',
      Agency: 'O',
      Version: 'D  96AA02031',
      TransactionSetID: 'CONTRL',
      Release: '0',
      Position: '1',
      SegmentID: 'UCI',
      Section: 'H',
      RequirementDesignator: 'M',
      MaximumLoopRepeat: '0',
      MaximumUsage: '1',
      Description: 'INTERCHANGE RESPONSE',
      LoopStart: true
    }
  };


// const data2={\"1\":{\"1\":{\"_id\":\"629f3e7a87b28e37a568dcc6\",\"Agency\":\"O\",\"Version\":\"D  96AA02031\",\"SegmentID\":\"UCI\",\"ElementID\":\"U020\",\"Release\":\"0\",\"Position\":\"1\",\"RequirementDesignator\":\"M\",\"SubElementReqDesignator\":\"M\",\"Type\":\"AN\",\"Description\":\"INTERCHANGE CONTROL REFERENCE\",\"MinimumLength\":\"1\",\"MaximumLength\":\"14\",\"CompositeElement\":\"FALSE\",\"RepeatFactor\":\"1\",\"SegmentPosition\":\"1\"},\"2\":{\"_id\":\"629f3e7a87b28e37a568dcc7\",\"Agency\":\"O\",\"Version\":\"D  96AA02031\",\"SegmentID\":\"UCI\",\"ElementID\":\"U004\",\"Release\":\"0\",\"Position\":\"2\",\"RequirementDesignator\":\"M\",\"GroupRequirementDesignatorID\":\"S002\",\"SubElementReqDesignator\":\"M\",\"GroupReqDesignator\":\"M\",\"GroupBeginEnd\":\"G\",\"Type\":\"AN\",\"Description\":\"SENDER / RECIVER IDENTIFICATION\",\"MinimumLength\":\"1\",\"MaximumLength\":\"35\",\"CompositeElement\":\"FALSE\",\"RepeatFactor\":\"1\",\"SegmentPosition\":\"1\"},\"3\":{\"_id\":\"629f3e7a87b28e37a568dcc8\",\"Agency\":\"O\",\"Version\":\"D  96AA02031\",\"SegmentID\":\"UCI\",\"ElementID\":\"U007\",\"Release\":\"0\",\"Position\":\"3\",\"RequirementDesignator\":\"M\",\"GroupRequirementDesignatorID\":\"S002\",\"SubElementReqDesignator\":\"O\",\"GroupReqDesignator\":\"M\",\"GroupBeginEnd\":\"R\",\"Type\":\"AN\",\"Description\":\"SENDER IDENTIFICATION QUALIFIER\",\"MinimumLength\":\"1\",\"MaximumLength\":\"4\",\"CompositeElement\":\"FALSE\",\"RepeatFactor\":\"0\",\"SegmentPosition\":\"1\"},\"4\":{\"_id\":\"629f3e7a87b28e37a568dcc9\",\"Agency\":\"O\",\"Version\":\"D  96AA02031\",\"SegmentID\":\"UCI\",\"ElementID\":\"U008\",\"Release\":\"0\",\"Position\":\"4\",\"RequirementDesignator\":\"M\",\"GroupRequirementDesignatorID\":\"S002\",\"SubElementReqDesignator\":\"O\",\"GroupReqDesignator\":\"M\",\"Type\":\"AN\",\"Description\":\"ADDRESS FOR REVERSE ROUNTING\",\"MinimumLength\":\"1\",\"MaximumLength\":\"14\",\"CompositeElement\":\"FALSE\",\"RepeatFactor\":\"0\",\"SegmentPosition\":\"1\"},\"5\":{\"_id\":\"629f3e7a87b28e37a568dcca\",\"Agency\":\"O\",\"Version\":\"D  96AA02031\",\"SegmentID\":\"UCI\",\"ElementID\":\"U010\",\"Release\":\"0\",\"Position\":\"5\",\"RequirementDesignator\":\"M\",\"GroupRequirementDesignatorID\":\"S003\",\"SubElementReqDesignator\":\"M\",\"GroupReqDesignator\":\"M\",\"GroupBeginEnd\":\"G\",\"Type\":\"AN\",\"Description\":\"RECIPIENT IDENTIFICATION\",\"MinimumLength\":\"1\",\"MaximumLength\":\"35\",\"CompositeElement\":\"FALSE\",\"RepeatFactor\":\"1\",\"SegmentPosition\":\"1\"},\"6\":{\"_id\":\"629f3e7a87b28e37a568dccb\",\"Agency\":\"O\",\"Version\":\"D  96AA02031\",\"SegmentID\":\"UCI\",\"ElementID\":\"U007\",\"Release\":\"0\",\"Position\":\"6\",\"RequirementDesignator\":\"M\",\"GroupRequirementDesignatorID\":\"S003\",\"SubElementReqDesignator\":\"O\",\"GroupReqDesignator\":\"M\",\"GroupBeginEnd\":\"R\",\"Type\":\"AN\",\"Description\":\"SENDER IDENTIFICATION QUALIFIER\",\"MinimumLength\":\"1\",\"MaximumLength\":\"4\",\"CompositeElement\":\"FALSE\",\"RepeatFactor\":\"0\",\"SegmentPosition\":\"1\"},\"7\":{\"_id\":\"629f3e7a87b28e37a568dccc\",\"Agency\":\"O\",\"Version\":\"D  9 6AA02031\",\"SegmentID\":\"UCI\",\"ElementID\":\"U014\",\"Release\":\"0\",\"Position\":\"7\",\"RequirementDesignator\":\"M\",\"GroupRequirementDesignatorID\":\"S003\",\"SubElementReqDesignator\":\"O\",\"GroupReqDesignator\":\"M\",\"Type\":\"AN\",\"Description\":\"ROUNTING ADDRESS\",\"MinimumLength\":\"1\",\"MaximumLength\":\"14\",\"CompositeElement\":\"FALSE\",\"RepeatFactor\":\"0\",\"SegmentPosition\":\"1\"},\"8\":{\"_id\":\"629f3e7a87b28e37a568dccd\",\"Agency\":\"O\",\"Version\":\"D  96AA02031\",\"SegmentID\":\"UCI\",\"ElementID\":\"U083\",\"Release\":\"0\",\"Position\":\"8\",\"RequirementDesignator\":\"M\",\"SubElementReqDesignator\":\"M\",\"Type\":\"ID\",\"Description\":\"ACTION, CODED\",\"MinimumLength\":\"1\",\"MaximumLength\":\"3\",\"CompositeElement\":\"FALSE\",\"RepeatFactor\":\"1\",\"SegmentPosition\":\"1\"}}}
const segmentData = Object.values(data);
let segmentDataHead=[];

// console.log(segmentDataHead);


for ( let x in segmentData){
    // console.log(x);
    console.log(segmentData[x]);
    for ( let y in segmentData[x]){
        console.log(y);
        segmentDataHead.push(y);
    }
}

// console.log(segmentData);

// there should be no agency , version and transactionsetid in the segment data

const table = new Table({
    // head: 
    //     new TableRow({
    //         children:segmentDataHead.map(function(cell) {
    //             return new TableCell({
    //                 children: [
    //                     new TextRun({
    //                         text: cell,
    //                         bold: true,
    //                     })
    //                 ]
    //             })
    //         }),
    //     })


    rows: [ new TableRow({
        children: [
          new TableCell({
            children: [
              new Paragraph({
                text: "POSITION",
                bold: true,
              }),
            ],
          }),
          new TableCell({
            children: [
              new Paragraph({
                text: "SEGMENT_ID",
                bold: true,
              }),
            ],
          }),
          new TableCell({
            children: [
              new Paragraph({
                text: "SEGMENT NAME",
                bold: true,
              }),
            ],
          }),
          new TableCell({
            children: [
              new Paragraph({
                text: "REQUEIRMENT NO",
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
        ],
    })





        segmentData.map(segment => {
        return new TableRow({   
            children: [
                new TableCell({
                  
                            children: [
                                new Paragraph({
                                    text: segment.Agency,
                                    bold: true
                                })]})
                                ,
                                new TableCell({
                  
                                    children: [
                                new Paragraph({
                                    text: segment.Version,
                                    bold: true
                                })]}),
                                new TableCell({
                  
                                    children: [
                                new Paragraph({
                                    text: segment.TransactionSetID,
                                    bold: true
                                })]}),
                                new TableCell({
                  
                                    children: [
                                new Paragraph({
                                    text: segment.Release,
                                    bold: true
                                })]}),
                                new TableCell({
                  
                                    children: [
                                new Paragraph({
                                    text: segment.Position,
                                    bold: true
                                })]}),
                                new TableCell({
                  
                                    children: [
                                new Paragraph({
                                    text: segment.SegmentID,
                                    bold: true
                                })]}),
                                new TableCell({
                  
                                    children: [
                                new Paragraph({
                                    text: segment.Section,
                                    bold: true
                                })]}),
                                new TableCell({
                  
                                    children: [
                                new Paragraph({
                                    text: segment.RequirementDesignator,
                                    bold: true
                                })]}),
                                new TableCell({
                  
                                    children: [
                                new Paragraph({
                                    text: segment.MaximumLoopRepeat,
                                    bold: true
                                })]}),
                                new TableCell({
                  
                                    children: [
                                new Paragraph({
                                    text: segment.MaximumUsage,
                                    bold: true
                                })]}),
                                new TableCell({
                  
                                    children: [
                                new Paragraph({
                                    text: segment.Description,
                                    bold: true
                                })]}),
                                new TableCell({
                  
                                    children: [
                                new Paragraph({
                                    text: segment.LoopStart,
                                    bold: true
                                })]}),
                            ]
                        })
    }),]
})] })

const doc = new Document({
    sections: [
        {
            children: [
                
                        table,
                        
                        
                    ],
              
            
        },
    ],

});




// const table = doc.createTable(10,10);
// table.getCell(1, 1).addContent(new docx.Paragraph("Hello"));


Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});

// Example of how you would create a table and add data to it
// Import from 'docx' rather than '../build' if you install from  Packer, Paragraph, Table, TableCell, TableRow, WidthType } from "../build";


