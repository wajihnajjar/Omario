const Excel = require("exceljs");
const https = require("https");
const fs = require("fs");
const { count } = require("console");
/*
Answers = [
    "Gafsa" , 
    "14-12-2022",
    "Omar Massoudi",
    "pc",
    "1500",
    "5",
    "Phone",
    "999",
    "10",
    

]
*/
let answers = [
    "Gafsa" , 
    "14-12-2022",
    "Omar Massoudi",
    "3",
    "pc",
    "1500",
    "5",
    "Phone",
    "999",
    "10",
]
const url = "https://res.cloudinary.com/dn6kxvylo/raw/upload/v1705007392/layo8xmv16sc5j5y20jd.xlsx"
const ProcessingFile = async (answers)=>{
        const workbook = new Excel.Workbook();
        await workbook.xlsx.readFile(`file.xlsx`).then(async () => {
        // H2 is the cell that have the place and date 
        // Asnwers[0] is the place and answers[1] is the date 
        workbook.worksheets[0].getCell("H2").value = answers[0]+','+answers[1]
        //C5 is the name and prenom 
        workbook.worksheets[0].getCell("C5").value = "Nom et Prenom : "+answers[2]
        //E10 is the cell that has the number of the facture 
        workbook.worksheets[0].getCell("E10").value="Quittance de commande N/"+answers[3]
        // Now Get the number of products we have 
        let NumberOfRowForProducts = (answers.length - 4)/3       
        // The products begin from Cell B13
         let PointerForNamesOfTheProducts = 4 
         let PointerForPriceOfTheProducts = 5     
         let PointerForQuantityOfTheProducts= 6
         let TotalSumOfTheProduct = 0 ; 
         // For loop for Begining from 1 to NumberOfRowProducts 
         for (let i = 1 ; i<=NumberOfRowForProducts ; i ++){
          let counter = (12+i).toString()
            workbook.worksheets[0].getCell('B'+counter).value = answers[PointerForNamesOfTheProducts]
            workbook.worksheets[0].getCell("E"+counter).value=answers[PointerForPriceOfTheProducts]
            workbook.worksheets[0].getCell("G"+counter).value=answers[PointerForQuantityOfTheProducts]
            let TotalOfThisProducts = parseInt(answers[PointerForPriceOfTheProducts])*parseInt(answers[PointerForQuantityOfTheProducts])
            TotalSumOfTheProduct+=TotalOfThisProducts
            workbook.worksheets[0].getCell("I"+counter).value = TotalOfThisProducts.toString()
            PointerForNamesOfTheProducts+=3 
            PointerForPriceOfTheProducts+=3 
            PointerForQuantityOfTheProducts+=3 


        }
        workbook.worksheets[0].getCell("H"+(12+NumberOfRowForProducts+2).toString()).value = "Total"
        workbook.worksheets[0].getCell("I"+(12+NumberOfRowForProducts+2).toString()).value = TotalSumOfTheProduct.toString()
        await workbook.xlsx.writeFile("output0.xlsx");
         })   
}
let MakeFacture = async (answers , url )=>{   
    const file = fs.createWriteStream("file.xlsx");
    https.get(url
    ,  (response)=>{
        response.pipe(file)
        file.on("finish",async ()=>
        {
            console.log("Downloading for the file Finish Now Processing it ")
            ProcessingFile(answers)
        })
    })
    return 

}
MakeFacture(answers,url)

