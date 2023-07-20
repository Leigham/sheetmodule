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
   *
   * @example
   * const credentials = {
   *   client_email: "your-client-email",
   *   private_key: "your-private-key",
   * };
   * const sheetManager = await SheetManager.getInstance(credentials);
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
   *
   * @example
   * await sheetManager.getAuth();
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
   *
   * @example
   * const sheetId = "your-spreadsheet-id";
   * const spreadsheetInfo = await sheetManager.getSheetInfo(sheetId);
   * console.log(spreadsheetInfo);
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
   *
   * @example
   * const sheetId = "your-spreadsheet-id";
   * const sheetIndex = 0; // Assuming you want the first sheet (0-based index)
   * const sheetValues = await sheetManager.getSheetValues(sheetId, sheetIndex);
   * console.log(sheetValues);
   */
  public async getSheetValues(
    id: string,
    sheetIndex: number
  ): Promise<sheets_v4.Schema$ValueRange | undefined> {
    const sheetName = await this.getSheetnameByIndex(id, sheetIndex);
    const res = await this.sheetClient?.spreadsheets.values.get({
      spreadsheetId: id,
      range: `${sheetName}!A1:Z`,
    });
    return res?.data;
  }

  /**
   * Gets the name of a specific sheet in a Google Sheets spreadsheet.
   * @param id The ID of the spreadsheet.
   * @param sheetIndex The index of the sheet in the spreadsheet (0-based).
   * @returns A Promise resolving to the name of the sheet.
   *
   * @example
   * const sheetId = "your-spreadsheet-id";
   * const sheetIndex = 0; // Assuming you want the first sheet (0-based index)
   * const sheetName = await sheetManager.getSheetnameByIndex(sheetId, sheetIndex);
   * console.log(sheetName);
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
   *
   * @example
   * const sheetId = "your-spreadsheet-id";
   * const sheetIndex = 0; // Assuming you want the first sheet (0-based index)
   * const columnToFilter = "Column A";
   * const filterValue = "Filter Value";
   * const filteredRow = await sheetManager.getSheetValuesByFilter(sheetId, sheetIndex, columnToFilter, filterValue);
   * console.log(filteredRow);
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
   *
   * @example
   * const sheetId = "your-spreadsheet-id";
   * const sheetIndex = 0; // Assuming you want the first sheet (0-based index)
   * const headers = await sheetManager.getSheetHeaders(sheetId, sheetIndex);
   * console.log(headers);
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
   *
   * @example
   * const sheetId = "your-spreadsheet-id";
   * const newSheetName = "New Sheet";
   * await sheetManager.addSheets(sheetId, newSheetName);
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
   *
   * @example
   * const sheetId = "your-spreadsheet-id";
   * const values = [
   *   {
   *     sheetName: "Sheet1",
   *     headers: ["Name", "Age", "Email"],
   *     rows: [
   *       ["John Doe", 30, "john@example.com"],
   *       ["Jane Smith", 25, "jane@example.com"],
   *     ],
   *   },
   *   {
   *     sheetName: "Sheet2",
   *     headers: ["Product", "Price", "Quantity"],
   *     rows: [
   *       ["Product A", 10.99, 50],
   *       ["Product B", 19.99, 30],
   *     ],
   *   },
   * ];
   * await sheetManager.addSheetValues(sheetId, values);
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
        await this.sheetClient?.spreadsheets.values.append({
          spreadsheetId: id,
          range: `${value.sheetName}!A1:Z`,
          valueInputOption: "RAW",
          requestBody: {
            values: [value.headers, ...value.rows],
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
   *
   * @example
   * const newSheetTitle = "New Spreadsheet";
   * const permissions = [
   *   {
   *     role: "writer",
   *     type: "user",
   *     emailAddress: "user@example.com",
   *   },
   *   {
   *     role: "reader",
   *     type: "domain",
   *     domain: "example.com",
   *   },
   * ];
   * const newDocument = await sheetManager.createNewDocument(newSheetTitle, permissions);
   * console.log(newDocument);
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
  public async getSheetURL(id: string): Promise<string | undefined | null> {
    const sheetInfo = await this.sheetClient?.spreadsheets.get({
      spreadsheetId: id,
    });
    return sheetInfo?.data.spreadsheetUrl;
  }
}

export default SheetManager;
