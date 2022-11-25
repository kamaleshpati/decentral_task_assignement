const ExcelJS = require('exceljs');
const fs = require('fs');

const wb = new ExcelJS.Workbook();
let obj = {}

if(process.argv.length<4){
    console.log("error: "+" less arguments")
}

else {
    let input = process.argv[2]
    let output = process.argv[3]

    if(!input.includes(".xlsx") || !output.includes(".json")) {
        console.log("wrong arguments");
    } else {
        wb.xlsx.readFile(input).then(() => {
    
            const ws = wb.getWorksheet(1);

            obj["name"] = ws.name;
            obj["styles"] = []
            obj["rows"] = {}
            
            ws.columns.forEach((col)=>{
                col.eachCell(c => {
                    let _index = -1
                    let fill = {};
                    fill["type"] = c.fill.type==undefined || c.fill.type==null || c.fill.type=="none"?"":c.fill.type
                    fill["pattern"] = c.fill.pattern==undefined || c.fill.pattern==null || c.fill.pattern=="none"?"":c.fill.pattern
                    fill["bgcolor"] = c.fill.bgColor==undefined || c.fill.bgColor==null || c.fill.bgColor=="none"?"":c.fill.bgColor.argb
                    for (let index = 0; index < obj["styles"].length; index++) {
                        const element = obj["styles"][index];
                        if(fill["type"] === element["type"] 
                                && fill["pattern"] === element["pattern"]
                                    && fill["bgcolor"] === element["bgcolor"]){
                            _index = index
                            break
                        }   
                    }
                    if(_index == -1){
                        obj["styles"].push(fill);
                        _index = obj["styles"].length-1
                    }
        
                    let val = "";
                    if(c.value == undefined || c.value==null)
                        val=""
                    else if(typeof c.value =='string')
                        val = c.value
                    else if(typeof c.value == 'object')
                        val = c.value.result
                    
                    if(c.row in obj["rows"]){
                        obj["rows"][c.row]["cells"][c.col] = {}
                        obj["rows"][c.row]["cells"][c.col]["text"]=val
                        obj["rows"][c.row]["cells"][c.col]["style"] = _index
                    } else {
                        obj["rows"][c.row] = {}
                        obj["rows"][c.row]["cells"] = {}
                        obj["rows"][c.row]["cells"][c.col] = {}
                        obj["rows"][c.row]["cells"][c.col]["text"]=val 
                        obj["rows"][c.row]["cells"][c.col]["style"] = _index
                    }
        
                });
            });
            
        
            fs.writeFile(output, JSON.stringify(obj), 'utf8', function (err) {
                if (err) {
                    console.log("An error occured while writing JSON Object to File.");
                    return console.log(err);
                }
             
                console.log("JSON file has been saved.");
            });
            
        }).catch(err => {
            console.log(err.message);
        });
        
    }
}

