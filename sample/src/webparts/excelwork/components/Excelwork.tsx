import * as React from 'react';
import styles from './Excelwork.module.scss';
import { IExcelworkProps } from './IExcelworkProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as ExcelJS from 'exceljs'
// import { sp } from '@pnp/sp';
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";
// import "@pnp/sp/site-users/web";
import { IFolder } from '@pnp/sp/folders';
import { saveAs } from "file-saver";
import { DefaultButton } from 'office-ui-fabric-react';
import { sp } from "@pnp/sp/presets/all";
import * as moment from 'moment';

export default class Excelwork extends React.Component<IExcelworkProps, {}> {
  public render(): React.ReactElement<IExcelworkProps> {
    return (
      <div className={styles.excelwork}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.description)}</p>
              {/* <input type='button' onClick={this.export} /> */}
              <DefaultButton text='export' onClick={this.export} />
            </div>
          </div>
        </div>
      </div>
    );
  }
  private async export() {

    let items = sp.web.lists.getByTitle("USER_ROLE").items.select("ID","OperatingCompany/ID").expand("OperatingCompany").getAll();
    console.log(items);
  }

  // private async export() {
  //   console.log('export');
  //   try {
  //     // get file
  //     const documentFolder: IFolder = sp.web.getFolderByServerRelativePath("Shared Documents");
  //     console.log(documentFolder)
  //     const fileBuffer: ArrayBuffer = await documentFolder.files.getByName("EI Billing Request.xlsx").getBuffer();
  //     // const fileBlob: Blob = await documentFolder.files.getByName("EI Billing Request.xlsx").getBlob();
  //     // console.log(fileBlob)

  //     // const excel: XLSX.WorkBook = XLSX.read(excelBuffer);
  //     // console.log(excel.SheetNames);
  //     // const ws: XLSX.WorkSheet = excel.Sheets["Raw Data"];
  //     // XLSX.utils.sheet_add_aoa(ws, heading);
  //     // XLSX.writeFileXLSX(excel, fileName.concat(this.fileExtension));

  //     // read file
  //     const wb = new ExcelJS.Workbook();
  //     // const a = wb.xlsx.read(fileBlob);
  //     await wb.xlsx.load(fileBuffer);
  //     // console.log(a);
  //     console.log(wb);
  //     console.log(wb.worksheets);

  //     let ws = await wb.getWorksheet("Raw Data");
  //     console.log(ws);
  //     // let aa = [3, "Sam", new Date()];
  //     // ws.addRow(aa).commit();

  //     const heading_allitems = [
  //       "Request No.",
  //       "Office",
  //       "Branch Office",
  //       "Currency",
  //       "Service Contract No.",
  //       "Repair Order Type",
  //       "Repair Order No.",
  //       "Work Description",
  //       "Repair Order Value",
  //       "Billing Type",
  //       "Percentage",
  //       "Invoice Amount",
  //       "Apply Sinking Fund / Rebate?",
  //       "Offset Amount",
  //       "Reviewer",
  //       "Sign-off By",
  //       "Created Date",
  //       "Created By",
  //       "Modified Date",
  //       "Modified By",
  //       "Submitted Date",
  //       "Submitted By",
  //       "Original Estimated Cost",
  //       "Repair Order Value",
  //       "Original Estimated C1%",
  //       "Actual Cost",
  //       "Committed / Late Cost",
  //       "Final Cost on Completion",
  //       "Repair Order Value",
  //       "Actual C1%",
  //       "Reviewed Date",
  //       "Reviewed By",
  //       "Signed-off Date",
  //       "Signed-off By",
  //       "Billing Info. Inputted Date",
  //       "Billing Info. Inputted By",
  //       "Billing Result",
  //       "Billing Document No.",
  //       "Billing Date",
  //       "SR Received?",
  //       "Payment Details Inputted Date",
  //       "Payment Details Inputted By",
  //       "Paid Date",
  //       "Status",
  //     ];

  //     ws.addRow(heading_allitems).commit();

  //     const buffer = await wb.xlsx.writeBuffer();
  //     const data = new Blob([buffer], { type: "application/octet-stream" });
  //     saveAs(data, "test.xlsx");

  //     // saveAs(fileBlob, 'test.xlsx');
  //   } catch (err) {
  //     console.log('error:');
  //     console.log(err);
  //   }
  // }
}

export interface AllItemsProperties {
  RequestNo: string;
  Office: string;
}
