import * as React from "react";
import styles from "./BulkUploadSpFx.module.scss";
import type { IBulkUploadSpFxProps } from "./IBulkUploadSpFxProps";
import * as XLSX from "xlsx";
import { getItemsFromList, saveData } from "./Services/Services";
import { getSP } from "./Services/pnpJs";
import {
  DefaultButton,
  PrimaryButton,
  ProgressIndicator,
} from "@fluentui/react";
import { convertDateOnlyToUTC } from "./Helper";

export interface ISharePointItem {
  Title: string;
  FirstName: string;
  LastName: string;
  WorkEmail: string;
  PersonalEmail: string;
  BirthDate: string | Date;
  HireDate: string | Date;
  WorkMode: string; // Consider using an enum if the choices are fixed
  Salary: number;
  IsMarried: boolean;
  SocialProfile: {
    Url: string;
  };
  JobTitle: string;
  About: string;
}

interface IMissingRow {
  rowNumber: number;
  missingFields: string[];
}

interface IState {
  listItems: ISharePointItem[];
  processedItems: ISharePointItem[]; // Stores valid items after file processing
  loading: boolean;
  error: string | null;
  isUploading: boolean;
  missingRows: IMissingRow[];
  percentComplete: number;
}

export default class BulkUploadSpFx extends React.Component<
  IBulkUploadSpFxProps,
  IState
> {
  constructor(props: IBulkUploadSpFxProps) {
    super(props);

    this.state = {
      listItems: [],
      processedItems: [],
      missingRows: [],
      isUploading: false,
      error: "",
      percentComplete: 0,
      loading: true,
    };
  }
  private fileInputRef = React.createRef<HTMLInputElement>();

  async getDataFromList() {
    this.setState({
      loading: true,
    });
    try {
      await getSP(this.context);
      const res = await getItemsFromList(
        this.props.context,
        "Employee Database"
      );
      this.setState({
        listItems: res,
        loading: false,
      });
    } catch (error) {
      console.error("Error fetching list items:", error);
      this.setState({
        listItems: [],
        loading: false,
      });
    }
  }

  handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files && event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt: ProgressEvent<FileReader>) => {
      const data = evt.target?.result;
      if (data) {
        const workbook = XLSX.read(data, { type: "array", cellDates: true });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData: any[] = XLSX.utils.sheet_to_json(worksheet, {
          raw: true,
        });

        // Array to store missing fields details for each row.
        const missingRowsData: IMissingRow[] = [];
        const itemsToUpload: ISharePointItem[] = [];

        jsonData.forEach((row, index) => {
          const missingFields: string[] = [];
          // Check each required field.
          if (!row.FirstName) missingFields.push("FirstName");
          if (!row.LastName) missingFields.push("LastName");
          if (!row.WorkEmail) missingFields.push("WorkEmail");
          if (!row.PersonalEmail) missingFields.push("PersonalEmail");
          if (!row.BirthDate) missingFields.push("BirthDate");
          if (!row.HireDate) missingFields.push("HireDate");
          if (!row.WorkMode) missingFields.push("WorkMode");
          if (!row.Salary) missingFields.push("Salary");
          if (!row.IsMarried) missingFields.push("IsMarried");
          if (!row.JobTitle) missingFields.push("JobTitle");
          if (!row.About) missingFields.push("About");
          if (!row.SocialProfile) missingFields.push("SocialProfile");

          // If missing fields exist, capture the row number (add 2 for header offset)
          if (missingFields.length > 0) {
            missingRowsData.push({
              rowNumber: index + 2,
              missingFields,
            });
          } else {
            const itemData: ISharePointItem = {
              Title: row.Title,
              FirstName: row.FirstName,
              LastName: row.LastName,
              WorkEmail: row.WorkEmail,
              PersonalEmail: row.PersonalEmail,
              // Ensure dates are in a valid format
              BirthDate: convertDateOnlyToUTC(row.BirthDate),
              HireDate: convertDateOnlyToUTC(row.HireDate),
              WorkMode: row.WorkMode,
              Salary: row.Salary,
              IsMarried: row.IsMarried === "Yes",
              SocialProfile: row.SocialProfile, // Expecting SocialProfile to be an object with Url
              JobTitle: row.JobTitle,
              About: row.About,
            };
            itemsToUpload.push(itemData);
          }
        });

        // If there are any missing fields, show the missing fields dialog.
        if (missingRowsData.length > 0) {
          this.setState({
            missingRows: missingRowsData,
            processedItems: [],
          });
          // Reset file input so user can re-select or correct the file.
          if (this.fileInputRef.current) {
            this.fileInputRef.current.value = "";
          }
          return;
        } else {
          // All rows are valid, save items in state and clear any previous missing rows.
          this.setState({
            missingRows: [],
            processedItems: itemsToUpload,
          });
          if (this.fileInputRef.current) {
            this.fileInputRef.current.value = "";
          }
        }
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Sequential upload: Uploads one item at a time.
  uploadItemsSequentially = async (itemsToUpload: ISharePointItem[]) => {
    const totalItems = itemsToUpload.length;
    for (let i = 0; i < totalItems; i++) {
      try {
        await saveData(
          itemsToUpload[i],
          this.props.context,
          "Employee Database"
        );
      } catch (error) {
        console.error(`Error uploading item at index ${i}`, error);
        // Optionally, decide whether to continue on error.
      }
      // Update progress (value between 0 and 1)
      this.setState({ percentComplete: (i + 1) / totalItems });
    }
    alert("Done");
  };

  // Parallel upload: Uploads all items concurrently.
  uploadItemsParallel = async (itemsToUpload: ISharePointItem[]) => {
    try {
      await Promise.all(
        itemsToUpload.map((item) =>
          saveData(item, this.props.context, "Employee Database").catch(
            (error) => {
              console.error("Error uploading item:", error);
              // Optionally return a value or error flag so that Promise.all doesn't reject.
            }
          )
        )
      );
      // For parallel upload, simply set progress to complete.
      this.setState({ percentComplete: 1 });
      alert("Done");
    } catch (error) {
      console.error("Error in parallel upload", error);
    }
  };

  // Handler for the "Upload sequentially" button.
  handleUploadSequential = async () => {
    const { processedItems } = this.state;
    if (processedItems.length === 0) return;
    this.setState({ isUploading: true, percentComplete: 0 });
    await this.uploadItemsSequentially(processedItems);
    this.setState({ isUploading: false, processedItems: [] });
    await this.getDataFromList();
  };

  // Handler for the "Upload parallely" button.
  handleUploadParallel = async () => {
    const { processedItems } = this.state;
    if (processedItems.length === 0) return;
    this.setState({ isUploading: true, percentComplete: 0 });
    await this.uploadItemsParallel(processedItems);
    this.setState({ isUploading: false, processedItems: [] });
    await this.getDataFromList();
  };

  public render(): React.ReactElement<IBulkUploadSpFxProps> {
    return (
      <section className={styles.bulkUploadSpFx}>
        <div className={styles.main}>
          <h1 className={styles.heading}>Bulk Upload SPFx</h1>
          <input
            type="file"
            accept=".xlsx, .xls"
            onChange={this.handleFileUpload}
            ref={this.fileInputRef}
          />
        </div>

        {/* Render upload buttons only if there are valid items and no upload is in progress */}
        {!this.state.isUploading && this.state.processedItems.length > 0 && (
          <div
            style={{
              display: "flex",
              gap: "12px",
              margin: "24px 0 0 0",
              justifyContent: "center",
              alignItems: "center",
            }}
          >
            <PrimaryButton
              text="Upload sequentially"
              onClick={this.handleUploadSequential}
            />
            <DefaultButton
              text="Upload parallely"
              onClick={this.handleUploadParallel}
            />
          </div>
        )}

        {this.state.missingRows.length > 0 && (
          <div className={styles.errorRows}>
            <p>
              The following rows have missing required fields. Please check your
              Excel file and upload again.
            </p>
            <ul>
              {this.state.missingRows.map((row) => (
                <li key={row.rowNumber}>
                  Row {row.rowNumber}: {row.missingFields.join(", ")}
                </li>
              ))}
            </ul>
          </div>
        )}

        <div>
          {this.state.isUploading && (
            <ProgressIndicator
              label="Uploading data..."
              description="Please wait while your file is being processed and uploaded. Do not close or refresh this window"
              percentComplete={this.state.percentComplete}
            />
          )}
        </div>
      </section>
    );
  }
}
