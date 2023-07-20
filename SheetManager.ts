import { google, drive_v3, sheets_v4 } from "googleapis";
import {
  Compute,
  Impersonated,
  JWT,
  UserRefreshClient,
  BaseExternalAccountClient,
} from "google-auth-library";

/**
 * SheetManager is a class that provides utility functions to interact with Google Sheets and Google Drive.
 */
class SheetManager {
  /**
   * The credentials to authenticate with Google APIs.
   */
  credentials: any; // Replace 'any' with the appropriate type for your credentials

  /**
   * The authentication instance to use for making API requests.
   * It can be an instance of `Compute`, `Impersonated`, `JWT`, `UserRefreshClient`, `BaseExternalAccountClient`, or `undefined`.
   */
  auth:
    | Compute
    | Impersonated
    | JWT
    | UserRefreshClient
    | BaseExternalAccountClient
    | undefined;

  /**
   * The instance of the Drive client.
   * This will be used to interact with the Google Drive API (v3).
   */
  driveClient: drive_v3.Drive | undefined;

  /**
   * The instance of the Sheets client.
   * This will be used to interact with the Google Sheets API (v4).
   */
  sheetClient: sheets_v4.Sheets | undefined;

  /**
   * Creates a new instance of SheetManager.
   * @param credentials The credentials used to authenticate with Google APIs.
   */
  private constructor(credentials: any) {
    this.credentials = credentials;
  }

  /**
   * Returns an instance of SheetManager.
   * This is an async function since it needs to initialize the authentication.
   * @param credentials The credentials used to authenticate with Google APIs.
   * @returns An instance of SheetManager.
   */
  public static async getInstance(credentials: any): Promise<SheetManager> {
    const sheetManager = new SheetManager(credentials);
    await sheetManager.getAuth();
    sheetManager.driveClient = google.drive({
      version: "v3",
      auth: sheetManager.auth,
    });
    sheetManager.sheetClient = google.sheets({
      version: "v4",
      auth: sheetManager.auth,
    });
    return sheetManager;
  }

  /**
   * Initializes the authentication instance for the SheetManager.
   * If the `auth` property is already set, this method does nothing.
   * @returns The authentication instance.
   * @throws An error if the authentication instance cannot be obtained.
   */
  public async getAuth(): Promise<
    Compute | Impersonated | JWT | UserRefreshClient | BaseExternalAccountClient
  > {
    if (!this.auth) {
      this.auth = await google.auth.getClient({
        credentials: this.credentials,
        scopes: [
          "https://www.googleapis.com/auth/drive",
          "https://www.googleapis.com/auth/spreadsheets",
        ],
      });
      if (!this.auth) throw new Error("Auth is undefined");
    }
    return this.auth;
  }

  /**
   * Gets information about a specific Google Sheets spreadsheet.
   * @param id The ID of the spreadsheet.
   * @returns A Promise resolving to the spreadsheet information.
   */
  public async getSheetInfo(
    id: string
  ): Promise<sheets_v4.Schema$Spreadsheet | undefined> {
    const res = await this.sheetClient?.spreadsheets.get({
      spreadsheetId: id,
    });
    return res?.data;
  }

  /**
   * Gets the values from a specific sheet in a Google Sheets spreadsheet.
   * @param id The ID of the spreadsheet.
   * @param sheetIndex The index of the sheet in the spreadsheet (0-based).
   * @returns A Promise resolving to the values from the sheet.
   */
  public async getSheetValues(
    id: string,
    sheetIndex: number
  ): Promise<sheets_v4.Schema$ValueRange | undefined> {
    const sheetName = await this.getSheetnameByIndex(id, sheetIndex);
    const res = await this.sheetClient?.spreadsheets.values.get({
      spreadsheetId: id,
      range: `${sheetName}!A1:Z`,
      valueRenderOption: "UNFORMATTED_VALUE",
    });
    return res?.data;
  }

  /**
   * Gets the name of a specific sheet in a Google Sheets spreadsheet.
   * @param id The ID of the spreadsheet.
   * @param sheetIndex The index of the sheet in the spreadsheet (0-based).
   * @returns A Promise resolving to the name of the sheet.
   */
  public async getSheetnameByIndex(
    id: string,
    sheetIndex: number
  ): Promise<string | undefined | null> {
    const sheetInfo = await this.sheetClient?.spreadsheets.get({
      spreadsheetId: id,
    });
    return sheetInfo?.data.sheets?.[sheetIndex].properties?.title;
  }

  /**
   * Gets the values from a specific row in a Google Sheets spreadsheet based on a filter value.
   * @param id The ID of the spreadsheet.
   * @param sheetIndex The index of the sheet in the spreadsheet (0-based).
   * @param col The column to filter on.
   * @param filter The value to filter by in the specified column.
   * @returns A Promise resolving to the values of the row that matches the filter.
   */
  public async getSheetValuesByFilter(
    id: string,
    sheetIndex: number,
    col: string,
    filter: string
  ): Promise<any[] | undefined | null> {
    const sheetName = await this.getSheetnameByIndex(id, sheetIndex);
    if (!sheetName) return undefined;

    // Get the column values
    const colRes = await this.sheetClient?.spreadsheets.values.get({
      spreadsheetId: id,
      range: `${sheetName}!${col}:${col}`,
      valueRenderOption: "UNFORMATTED_VALUE",
    });
    const colValues = colRes?.data.values;

    if (!colValues) return [];

    // Find the row that matches the filter
    const row = colValues.findIndex((row) => row[0] === filter);
    if (row === -1) return [];

    // Get the entire row using the found index
    const rowRes = await this.sheetClient?.spreadsheets.values.get({
      spreadsheetId: id,
      range: `${sheetName}!A${row + 1}:Z${row + 1}`,
    });
    return rowRes?.data.values;
  }

  /**
   * Gets the headers (first row) from a specific sheet in a Google Sheets spreadsheet.
   * @param id The ID of the spreadsheet.
   * @param sheetIndex The index of the sheet in the spreadsheet (0-based).
   * @returns A Promise resolving to the headers of the sheet.
   */
  public async getSheetHeaders(
    id: string,
    sheetIndex: number
  ): Promise<string[] | undefined> {
    const sheetName = await this.getSheetnameByIndex(id, sheetIndex);
    if (!sheetName) return undefined;

    const res = await this.sheetClient?.spreadsheets.values.get({
      spreadsheetId: id,
      range: `${sheetName}!1:1`,
    });
    return res?.data.values?.[0] ?? [];
  }

  /**
   * Adds a new sheet to a Google Sheets spreadsheet.
   * @param id The ID of the spreadsheet.
   * @param sheetName The name of the new sheet to add.
   * @returns A Promise resolving when the new sheet is added.
   */
  private async addSheets(id: string, sheetName: string): Promise<void> {
    const resp = await this.sheetClient?.spreadsheets.get({
      spreadsheetId: id,
    });
    const sheets = resp?.data.sheets;
    if (!sheets) throw new Error("Sheets is undefined");
    const sheetExists = sheets.some(
      (sheet) => sheet.properties?.title === sheetName
    );
    if (!sheetExists) {
      await this.sheetClient?.spreadsheets.batchUpdate({
        spreadsheetId: id,
        requestBody: {
          requests: [
            {
              addSheet: {
                properties: {
                  title: sheetName,
                },
              },
            },
          ],
        },
      });
    }
  }

  /**
   * Adds new values to a specific sheet in a Google Sheets spreadsheet.
   * @param id The ID of the spreadsheet.
   * @param values An array of objects containing the sheet name, headers, and rows to add.
   * @returns A Promise resolving when the values are added.
   */
  public async addSheetValues(
    id: string,
    values: {
      sheetName: string;
      headers: string[];
      rows: (string | number | boolean)[][];
    }[]
  ): Promise<void> {
    await Promise.all(
      values.map(async (value) => {
        await this.addSheets(id, value.sheetName);

        // Calculate the dimensions of the new data
        const numRows = value.rows.length;
        const numColumns = value.headers.length;

        // Get the existing sheet properties
        const sheetProperties = await this.sheetClient?.spreadsheets.get({
          spreadsheetId: id,
          ranges: [`${value.sheetName}!A1`],
          fields: "sheets(properties)",
        });
        const sheets = sheetProperties?.data.sheets;
        if (!sheets) throw new Error("Sheets is undefined");
        const sheetindex = sheets.findIndex((sheet) => {
          return sheet.properties?.title === value.sheetName;
        });
        if (sheetindex === -1) throw new Error("Sheet index is -1");
        const sheetId = sheets[sheetindex].properties?.sheetId;
        if (!sheetId) throw new Error("Sheet ID is undefined");

        // Get the existing grid properties
        const gridProperties =
          sheetProperties?.data.sheets?.[sheetindex]?.properties
            ?.gridProperties;

        if (!gridProperties) {
          throw new Error("Failed to get grid properties");
        }

        // Update the grid properties to match the new dimensions
        gridProperties.rowCount = Math.max(
          gridProperties.rowCount || 0,
          numRows + 1 // Add 1 to account for the header row
        );
        gridProperties.columnCount = Math.max(
          gridProperties.columnCount || 0,
          numColumns
        );

        // Update the sheet properties with the new grid properties
        await this.sheetClient?.spreadsheets.batchUpdate({
          spreadsheetId: id,
          requestBody: {
            requests: [
              {
                updateSheetProperties: {
                  properties: {
                    title: value.sheetName,
                    sheetId: sheetId,
                    gridProperties: {
                      ...gridProperties,
                      rowCount: value.rows.length + 1,
                      columnCount: value.headers.length,
                    },
                  },
                  fields: "gridProperties",
                },
              },
            ],
          },
        });

        // Create data validation requests based on the header row
        const requests = value.headers.map((header, index) => {
          const dataType = typeof value.rows[0][index];
          const rule: any = {
            condition: {
              type: "CUSTOM_FORMULA",
            },
            showCustomUi: true,
          };

          let validationFormula: string;
          const columnLetter = String.fromCharCode(65 + index);
          if (dataType === "string") {
            validationFormula = `=ISTEXT(${columnLetter}2:${columnLetter})`;
          } else if (dataType === "number") {
            validationFormula = `=ISNUMBER(${columnLetter}2:${columnLetter})`;
          } else if (dataType === "boolean") {
            validationFormula = `=OR(${columnLetter}2:${columnLetter}=TRUE, ${columnLetter}2:${columnLetter}=FALSE)`;
          } else {
            validationFormula = "";
          }

          if (validationFormula) {
            rule.condition.values = [{ userEnteredValue: validationFormula }];
          }
          return {
            setDataValidation: {
              range: {
                sheetId: sheetId,
                startRowIndex: 1,
                endRowIndex: numRows + 1,
                startColumnIndex: index,
                endColumnIndex: index + 1,
              },
              rule: rule,
            },
          };
        });

        // Append the values to the sheet
        await this.sheetClient?.spreadsheets.values.append({
          spreadsheetId: id,
          range: `${value.sheetName}!A1:Z`,
          valueInputOption: "RAW",
          requestBody: {
            values: [value.headers, ...value.rows],
          },
        });
        await this.sheetClient?.spreadsheets.batchUpdate({
          spreadsheetId: id,
          requestBody: {
            requests: requests,
          },
        });
      })
    );
  }

  /**
   * Creates a new Google Sheets document with the given title and permissions.
   * @param title The title of the new document.
   * @param permissions An array of permission objects to apply to the document.
   * @returns A Promise resolving to the created document information.
   * @throws An error if the document ID is undefined.
   */
  public async createNewDocument(
    title: string,
    permissions: drive_v3.Schema$Permission[]
  ): Promise<drive_v3.Schema$File> {
    const newSheetDocument = await this.driveClient?.files.create({
      requestBody: {
        name: title,
        mimeType: "application/vnd.google-apps.spreadsheet",
      },
    });
    if (!newSheetDocument?.data.id) throw new Error("Document id is undefined");
    await Promise.all(
      permissions.map(async (permission) => {
        await this.driveClient?.permissions.create({
          fileId: newSheetDocument?.data.id!,
          requestBody: permission,
          transferOwnership: permission.role === "owner",
          sendNotificationEmail: false,
        });
      })
    );
    return newSheetDocument?.data;
  }

  /**
   * Gets the URL of a specific Google Sheets spreadsheet.
   * @param id The ID of the spreadsheet.
   * @returns A Promise resolving to the URL of the spreadsheet.
   */
  public async getSheetURL(id: string): Promise<string | undefined | null> {
    const sheetInfo = await this.sheetClient?.spreadsheets.get({
      spreadsheetId: id,
    });
    return sheetInfo?.data.spreadsheetUrl;
  }
}

export default SheetManager;
