// import React from 'react'
// const ExcelJS = require('exceljs');
// interface IProps {
//   history:string[],
// }
// interface IState {
//   items:Array<any>,
// }

// class App extends React.Component<IProps, IState> {
//   constructor(props: any) {
//     super(props);
//     this.state = { items: [] }
//   }  
//   componentDidMount(){
//     const self = this;
//     window.document.getElementById("file1").addEventListener("change", function() {
//       console.log("#change");
//       self.uploadFile();
//     });    
//   }
//   uploadFile(){
//     const self = this;
//     console.log("uploadFile");
//     const files = document.querySelector<HTMLInputElement>('#file1').files;
//     const fileObject = files[0]; 
//     if (typeof fileObject === "undefined") {
//       console.error("none, fileObject");
//       return;
//     }
// //console.log(fileObject);
//     const blobURL = window.URL.createObjectURL(fileObject);
//     console.log(blobURL);
//     const xhr = new XMLHttpRequest();
//     xhr.onload = function() {
//       const result = xhr.response; // ArrayBuffer
// //      console.log(result);
//       const data = new Uint8Array(result);
// //      console.log(data);
//       self.loadExcelData(data);
//     };
//     xhr.responseType = "arraybuffer";
//     xhr.open("GET", blobURL);
//     xhr.send();    
//     console.log("start-upload");
//   }
//   async loadExcelData(data: any){
//     try{
//       const workbook = new ExcelJS.Workbook();
//       await workbook.xlsx.load(data);
//       const worksheet = workbook.getWorksheet('sheet1');
//       worksheet.pageSetup = {orientation:'portrait'};
//       const startRow = 4;
//       const items: any[] = [];
//       let row = worksheet.getRow(1);
//       for (let i = startRow; i < 11; i++) {
//         row = worksheet.getRow(i);
//         if(row.getCell(1).value !== null){
//           console.log(row.getCell(1).value);
//           let item = {
//             ID: row.getCell(1).value,
//             NAME: row.getCell(2).value,
//             VALUE: row.getCell(3).value,
//           }
//           items.push(item);
//         }
//       }
//   //    console.log(items);
//       this.setState({items:items });
//       alert("complete load data");
//     } catch (e) {
//       console.error(e);
//       alert("Error, load data");
//     }    
//   }
//   render(){
//     console.log(this.state);
//     return(
//     <div className="container">
//       <h1>xls7: read Sample</h1>
//       <hr />
//       File: <br />
//       <input type="file" name="file1" id="file1" /><br />
//       <hr className="my-1" />
//       {/* table */}
//       <h3>XXX経費資料</h3>
//       <table className='table'>
//         <thead>
//         <tr>
//           <th>ID</th>
//           <th>NAME</th>
//           <th>Price</th>
//         </tr>
//         </thead>
//         <tbody>
//         { this.state.items.map((item: any, index: number) => (
//           <tr key={index}>
//             <td>{item.ID}</td>
//             <td>{item.NAME}</td>
//             <td>{item.VALUE} JPY</td>
//           </tr>
//         ))        
//         }
//         </tbody>
//       </table>
//     </div>
//     )
//   }
// }

// export default App;