 const xlsx = require('xlsx')

 const buscaEmail = async () => {

   let emailBD = []
   let emailTerceiros = [];
   let mergedEmail = [];
   let holder = [];
   
   try {
     let workbook = xlsx.readFile('./emailDb.xlsx');
     
     const worksheet = workbook.Sheets[workbook.SheetNames[0]];
     
     
     //Busca os valores recebidos do banco de dados
     const fromDB = (worksheet) => {
       const columnA = [];

       for (let row in worksheet) {
         if (row.toString()[0] === 'A') {
           columnA.push({localDb:worksheet[row].v});
         }
       }
       
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

       //Adiciona os emails de terceiros
       emailTerceiros.forEach((data) => {
         let email = data.thirdEmail
         if (email.includes(atEmail) === true) {
           holder.push(email);
         }
       })

       holder.push(data.localDb);
     });
    //  Cria um novo array Sem a repetição
     let noRepeat = [ ...new Set(holder) ];
     
    
     noRepeat.forEach((data)=>{
       mergedEmail.push({mergedEmail: data})
     });
     console.log(mergedEmail);
     let newWB = xlsx.utils.book_new();
     let newWS = xlsx.utils.json_to_sheet(mergedEmail);
    
     xlsx.utils.book_append_sheet(newWB,newWS,'mergedData');
     
     
     xlsx.writeFile(newWB,'emailDbMerged.xlsx')

   } catch (err) {
     console.log(err);
   }

 }

 buscaEmail();