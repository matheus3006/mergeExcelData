 const xlsx = require('xlsx')

 const buscaEmail = async () => {

   let emailBD = []
   let emailTerceiros = [];
   let mergedEmail = [];
   try {
     const workbook = xlsx.readFile('./emailDb.xlsx');
     const worksheet = workbook.Sheets[workbook.SheetNames[0]];

     //Busca os valores recebidos do banco de dados
     const fromDB = (worksheet) => {
       const columnA = [];

       for (let row in worksheet) {
         if (row.toString()[0] === 'A') {
           columnA.push({localDb:worksheet[row].v});
         }
       }
       columnA.shift();

       return columnA;
     };

     //  Busca os valores adicionais 
     const fromThird = (worksheet) => {
       const columnB = [];

       for (let row in worksheet) {
         if (row.toString()[0] === 'B') {
           columnB.push({thirdEmail:worksheet[row].v})
         }
       }
       columnB.shift();
       return columnB;
     }

     emailBD = fromDB(worksheet)
     emailTerceiros = fromThird(worksheet)
     
    


     emailBD.forEach((data) => {
      //Cria função na qual vai buscar os valores de email
       const findAtEmail = (data) => {
         let email = data.localDb;
         let holder = email.split('@');
         return holder[1];
       }
       //Chama a função
       let atEmail = findAtEmail(data);


       emailTerceiros.forEach((data) => {
         let email = data.thirdEmail
         if (email.includes(atEmail) === true) {
           mergedEmail.push({mergedEmail: email});
         }
       })

       mergedEmail.push({mergedEmail: data.localDb});
     });
     
    

     let newWB = xlsx.utils.book_new();
     let newWS = xlsx.utils.json_to_sheet(mergedEmail);
     console.log(newWS);
     xlsx.utils.book_append_sheet(newWB,newWS,'mergedData');
     
     xlsx.writeFile(newWB,'emailDB.xlsx')

   } catch (err) {
     console.log(err);
   }

 }

 buscaEmail();