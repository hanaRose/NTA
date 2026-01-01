// src/extensions/shipWork/ShipWorkCommandSet.ts

import { override } from '@microsoft/decorators';
import {
  BaseListViewCommandSet,
  IListViewCommandSetExecuteEventParameters,
  IListViewCommandSetListViewUpdatedParameters,
  RowAccessor
} from '@microsoft/sp-listview-extensibility';
  import { SPHttpClient } from '@microsoft/sp-http'; // ודאי שיש לך את זה למעלה בקובץ


import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  Header,
  Footer,
  ImageRun
} from 'docx';

import { saveAs } from 'file-saver';
//import { SPHttpClient } from '@microsoft/sp-http';

// תמונות להידר / פוטר – קבועים (עובד לכל משתמש)
const HEADER_IMG_URL: string =
  'https://ntaisrael.sharepoint.com/sites/M365_TenderProposals/SiteAssets/AgendaHeaderFooter/header.jpg';

const FOOTER_IMG_URL: string =
  'https://ntaisrael.sharepoint.com/sites/M365_TenderProposals/SiteAssets/AgendaHeaderFooter/footer.jpg';

export interface IShipWorkCommandSetProperties {}

export default class ShipWorkCommandSet
  extends BaseListViewCommandSet<IShipWorkCommandSetProperties> {

  @override
  public async onInit(): Promise<void> {
    console.log('ShipWorkCommandSet initialized');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const cmd1 = this.tryGetCommand('COMMAND_1');

    const hasSelection = !!event.selectedRows && event.selectedRows.length > 0;

    if (cmd1) {
      cmd1.visible = hasSelection; // כפתור יצוא – רק אם יש בחירה
    }
  }
  /*
  @override
  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    switch (event.itemId) {
      case 'COMMAND_1': {
        const rows = event.selectedRows;
        if (!rows || rows.length === 0) {
          alert('Please select at least one item to export.');
          return;
        }

        try {
          console.log('PMO rows selected:', rows);

          // 1. טוענים את כל הפריטים מרשימת Integration לפי ה-lookup בכל שורה של PMO decisions
          const integrationItems = await Promise.all(
            rows.map(r => this._fetchIntegrationItemForRow(r))
          );

          console.log('Loaded Integration items array:', integrationItems);

          // 2. טוענים את התמונות כ-ArrayBuffer
          const headerImg = await this._loadImage(HEADER_IMG_URL);
          const footerImg = await this._loadImage(FOOTER_IMG_URL);

          // 3. בונים את המסמך משילוב PMO + Integration לכל פריט
          const doc = this._buildAgendaDocument(rows, integrationItems, headerImg, footerImg);
          const blob = await Packer.toBlob(doc);
          saveAs(blob, 'Integration_Agenda.docx');
        } catch (e) {
          console.error('Failed to build agenda document', e);
          alert('Error while creating the agenda document. Check console for details.');
        }
        break;
      }

      default:
        break;
    }
  }*/

@override
public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
  switch (event.itemId) {
    case 'COMMAND_1': {
      const rows = event.selectedRows;
      if (!rows || rows.length === 0) {
        alert('Please select at least one item to export.');
        return;
      }

      try {
        console.log('Integration rows selected:', rows);

        // 1) מביאים FULL ITEM לכל שורה שנבחרה ב-Integration (כולל כל העמודות)
        const integrationItems = await Promise.all(
          rows.map(r => this._fetchIntegrationItemForRow(r))
        );

        console.log('😑 Loaded Integration items array:', integrationItems);

        // 2) תמונות
        const headerImg = await this._loadImage(HEADER_IMG_URL);
        const footerImg = await this._loadImage(FOOTER_IMG_URL);

        // 3) בונים מסמך. (rows = שורות Integration, integrationItems = פריטים מלאים)
        const doc = this._buildAgendaDocument(rows, integrationItems, headerImg, footerImg);

        const blob = await Packer.toBlob(doc);
        saveAs(blob, 'Integration_Agenda.docx');
      } catch (e) {
        console.error('Failed to build agenda document', e);
        alert('Error while creating the agenda document. Check console for details.');
      }
      break;
    }

    default:
      break;
  }
}

  // ================== קריאה לרשימת Integration ==================
/*
  private async _fetchIntegrationItemForRow(row: RowAccessor): Promise<any> {
  try {
    const lookup = row.getValueByName("Integration");
    if (!lookup || !Array.isArray(lookup) || lookup.length === 0) {
      console.warn("No lookup value found for Integration field");
      return null;
    }

    const lookupId = lookup[0].lookupId;
    if (!lookupId) {
      console.warn("Integration lookup has no ID");
      return null;
    }

    console.log("Fetching Integration item ID:", lookupId);

    const guid = "2c962132-409d-4bf2-9440-3b3b6c7975a0"; // GUID של הרשימה
    const url =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists(guid'${guid}')/items(${lookupId})?$select=*`;

    const response = await fetch(url, {
      headers: {
        Accept: "application/json;odata=nometadata"
      }
    });

    if (!response.ok) {
      console.error("Failed to fetch Integration item", lookupId, response.status);
      return null;
    }

    const item = await response.json();
    console.log("Integration item fetched:", item);
    return item;

  } catch (e) {
    console.error("Error fetching Integration item", e);
    return null;
  }
}
*/
/*
private async _fetchIntegrationItemForRow(row: RowAccessor): Promise<any | null> {
  try {
    console.log("😕 row ", row);
    // ברשימת Integration יש תמיד עמודת ID
    const idRaw: any = row.getValueByName('ID') ?? row.getValueByName('Id');
    console.log("🤫idRaw ", idRaw)
    const itemId = Number(idRaw);
    console.log("itemId ", itemId);

    
    if (!itemId || Number.isNaN(itemId)) {
      console.warn('Could not resolve Integration item ID from selected row', { idRaw });
      return null;
    }

    console.log('Fetching Integration item ID:', itemId);

    const guid = '2c962132-409d-4bf2-9440-3b3b6c7975a0'; // GUID של Integration
    const url =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/_api/web/lists(guid'${guid}')/items(${itemId})?$select=*`;

    const resp = await this.context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );

    if (!resp.ok) {
      console.error('Failed to fetch Integration item', itemId, resp.status);
      return null;
    }

    const item = await resp.json();
    console.log('Integration item fetched:', item);
    return item;
  } catch (e) {
    console.error('Error fetching Integration item', e);
    return null;
  }
}
*/

  private async _fetchIntegrationItemForRow(row: RowAccessor): Promise<any | null> {
    try {
      const idRaw: any = row.getValueByName('ID') ?? row.getValueByName('Id');
      const itemId = Number(idRaw);

      if (!itemId || Number.isNaN(itemId)) {
        console.warn('Could not resolve Integration item ID from selected row', { idRaw });
        return null;
      }

      const guid = '2c962132-409d-4bf2-9440-3b3b6c7975a0';
      const url =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists(guid'${guid}')/items(${itemId})?$select=*&$format=json`; // <-- כופה JSON

      const resp = await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata', // <-- זה ה-SharePoint "הקלאסי"
            'odata-version': ''                         // <-- מונע בעיות OData
          }
        }
      );

      const contentType = resp.headers.get('content-type') || '';
      const bodyText = await resp.text(); // קוראים טקסט כדי לא להיתקע על json()

      if (!resp.ok) {
        console.error('Failed to fetch Integration item', itemId, resp.status, bodyText.slice(0, 500));
        return null;
      }

      // אם בכל זאת הגיע XML/HTML – נדפיס ונחזור null (כדי שתראי מה הוחזר)
      if (!contentType.toLowerCase().includes('application/json')) {
        console.error(
          'Expected JSON but got:',
          { itemId, contentType, preview: bodyText.slice(0, 500) }
        );
        return null;
      }

      const item = JSON.parse(bodyText);
      console.log('Integration item fetched:', item);
      return item;

    } catch (e) {
      console.error('Error fetching Integration item', e);
      return null;
    }
  }


  // ================== בניית מסמך ה-Word ==================

  private _buildAgendaDocument(
    rows: ReadonlyArray<RowAccessor>,
    integrationItems: (any | null)[],
    headerImg?: ArrayBuffer,
    footerImg?: ArrayBuffer
  ): Document {
    const firstRow = rows[0];
    const firstIntegration = integrationItems[0];

    // HEADER – קודם PMO, אם ריק – מ-Integration
    const meetingDate =
      this._getTextFromRow(firstRow, 'ApplicationDate') ||
      this._getTextFromIntegration(firstIntegration, 'ApplicationDate');

    const originatingLM =
      this._getTextFromRow(firstRow, 'OriginatingLineManager') ||
      this._getTextFromIntegration(firstIntegration, 'OriginatingLineManager');

    const tenderNumber =
      this._getTextFromRow(firstRow, 'TenderNumber') ||
      this._getTextFromIntegration(firstIntegration, 'TenderNumber');

    const tenderPhase =
      this._getTextFromRow(firstRow, 'TenderPhase') ||
      this._getTextFromIntegration(firstIntegration, 'TenderPhase');

    const children: (Paragraph | Table)[] = [];

    children.push(
      new Paragraph({
        text: 'Integration Team Meeting Agenda',
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.LEFT,
        bidirectional: false
      })
    );

    children.push(new Paragraph({ text: '', alignment: AlignmentType.LEFT,bidirectional: false }));

    children.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        bidirectional: false,
        children: [
          new TextRun({ text: 'Meeting date: ', bold: true }),
          new TextRun(meetingDate || '[dd/mm/yyyy]')
        ]
      })
    );

    children.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        bidirectional: false,
        children: [
          new TextRun({ text: 'Originating Line Manager: ', bold: true }),
          new TextRun(originatingLM || '[M1/M2/M3]')
        ]
      })
    );

    children.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        bidirectional: false,
        children: [
          new TextRun({ text: 'Tender No.: ', bold: true }),
          new TextRun(tenderNumber || '[Tender number]')
        ]
      })
    );

    children.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        bidirectional: false,
        children: [
          new TextRun({ text: 'Tender Phase: ', bold: true }),
          new TextRun(tenderPhase || '[XXXXX]')
        ]
      })
    );

    children.push(new Paragraph({ text: '', alignment: AlignmentType.LEFT, bidirectional: false }));

    // לכל פריט – טבלה לפי התבנית, שילוב PMO + Integration
    rows.forEach((row, index) => {
      const integ = integrationItems[index];
      console.log("😁integ ", integ);
      const number = index + 1;
      children.push(...this._buildProposedChangeSection(row, integ, number));
    });

    children.push(
      new Paragraph({
        alignment: AlignmentType.LEFT,
        bidirectional: false,
        spacing: { before: 400 },
        children: [
          new TextRun({
            text:
              '[Additional LMs / Tenders / proposed changes will be added as applicable]',
            italics: true
          })
        ]
      })
    );

    const header = new Header({
      children: [
        headerImg
          ? new Paragraph({
              alignment: AlignmentType.LEFT,
              bidirectional: false,
              children: [
                new ImageRun({
                  data: headerImg,
                  type: 'image/jpeg',
                  transformation: {
                    width: 700,
                    height: 80
                  }
                } as any)
              ]
            })
          : new Paragraph({
              alignment: AlignmentType.LEFT,
              bidirectional: false,
              children: [
                new TextRun({ text: 'Integration Team Meeting Agenda', bold: true })
              ]
            })
      ]
    });

    const footer = new Footer({
      children: [
        footerImg
          ? new Paragraph({
              alignment: AlignmentType.LEFT,
              bidirectional: false,
              children: [
                new ImageRun({
                  data: footerImg,
                  type: 'image/jpeg',
                  transformation: {
                    width: 700,
                    height: 50
                  }
                } as any)
              ]
            })
          : new Paragraph({
              alignment: AlignmentType.LEFT,
              bidirectional: false,
              children: [
                new TextRun({
                  text: 'NTA – Integration Team | Confidential',
                  size: 18
                })
              ]
            })
      ]
    });

    return new Document({
      sections: [
        {
          properties: {
            page: {
              margin: { top: 720, bottom: 720, left: 720, right: 720 }
            }
          },
          headers: { default: header },
          footers: { default: footer },
          children
        }
      ]
    });
  }

  // ====== בלוק של Proposed Change – טבלה במבנה התבנית ======

  private _buildProposedChangeSection(
    pmoRow: RowAccessor,
    integrationItem: any | null,
    number: number
  ): (Paragraph | Table)[] {
    const result: (Paragraph | Table)[] = [];

    result.push(
      new Paragraph({
        text: `Proposed Change #${number}`,
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 400, after: 200 },
        alignment: AlignmentType.LEFT,
        bidirectional: false
      })
    );

    // קודם PMO, אם ריק – Integration

    const ntaRef =
      this._getTextFromRow(pmoRow, 'NTA_x2019_s_x0020_reference') ||
      this._getTextFromIntegration(integrationItem, 'NTA_x2019_s_x0020_reference');

    const part =
      this._getTextFromRow(pmoRow, 'Part') ||
      this._getTextFromIntegration(integrationItem, 'Part');

    const volume =
      this._getTextFromRow(pmoRow, 'Volume') ||
      this._getTextFromIntegration(integrationItem, 'Volume');

    const chapter =
      this._getTextFromRow(pmoRow, 'Chapter') ||
      this._getTextFromIntegration(integrationItem, 'Chapter');

    const section =
      this._getTextFromRow(pmoRow, 'SectionNumber') ||
      this._getTextFromIntegration(integrationItem, 'SectionNumber');

    const category =
      this._getTextFromRow(pmoRow, 'Category') ||
      this._getTextFromIntegration(integrationItem, 'Category');

    const decision =
      this._getTextFromRow(pmoRow, 'DecisionAppliesToOtherWorksTende') ||
      this._getTextFromIntegration(integrationItem, 'DecisionappliestootherWorksTende');

    const particulars =
      this._getTextFromRow(pmoRow, 'ParticularsAndDescriptionOfPropo') ||
      this._getTextFromIntegration(integrationItem, 'ParticularsAndDescriptionOfPropo');

    const explanation =
      this._getTextFromRow(pmoRow, 'Explanation_x0028_s_x0029_Of_x00') ||
      this._getTextFromIntegration(integrationItem, 'Explanation_x0028_s_x0029_Of_x00');

    const existingPmo = this._getTextFromRow(pmoRow, 'Existingwordingofapplicableprovi');

    const existingInt1 = this._getTextFromIntegration(integrationItem, 'Existingwordingofapplicableprovi');
    const existingInt2 = this._getTextFromIntegration(integrationItem, 'ExistingwordingofapplicableRFCre');
    console.log("😎🤩 existingInt1 Existingwordingofapplicableprovi - ", existingInt1, " integration item ", integrationItem);
    console.log("😎 existingInt2 ExistingwordingofapplicableRFCre  - ", existingInt2);
    const existing = existingPmo || existingInt1 || existingInt2;

    const proposedPmo = this._getTextFromRow(pmoRow, 'Proposedrevisions_x0028_s_x0029_');
    const proposedInt1 = this._getTextFromIntegration(integrationItem, 'Proposedrevisions_x0028_s_x0029_');
    const proposedInt2 = this._getTextFromIntegration(integrationItem, 'Proposedrevisions_x0028_s_x0029_0');
    const proposed = proposedPmo || proposedInt1 || proposedInt2;

    const recPmo = this._getTextFromRow(pmoRow, 'Recommendation_x0028_s_x0029_wit');
    const recInt1 = this._getTextFromIntegration(integrationItem, 'Recommendation_x0028_s_x0029_wit');
    const recInt2 = this._getTextFromIntegration(integrationItem, 'Recommendation_x0028_s_x0029_wit0');
    const recommendation = recPmo || recInt1 || recInt2;

    const table = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            this._headerCell("NTA's reference #"),
            this._headerCell('Part'),
            this._headerCell('Volume'),
            this._headerCell('Chapter'),
            this._headerCell('Section'),
            this._headerCell('Category'),
            this._headerCell('Decision applies to other Works Tenders?')
          ]
        }),
        new TableRow({
          children: [
            this._valueCell(ntaRef || '[XXXXX]'),
            this._valueCell(part || '[XXXXX]'),
            this._valueCell(volume || '[XXXXX]'),
            this._valueCell(chapter || '[XXXXX]'),
            this._valueCell(section || '[XXXXX]'),
            this._valueCell(category || '[Technical / Legal / Financial]'),
            this._valueCell(decision || '[XXX]')
          ]
        }),
        new TableRow({
          children: [
            this._labelCell('Particulars and description\nof Proposed Change'),
            this._valueCell(particulars || '[XXXXX]', 6)
          ]
        }),
        new TableRow({
          children: [
            this._labelCell('Explanation(s) of, and\nreason(s) for, Proposed\nChange'),
            this._valueCell(explanation || '[XXXXX]', 6)
          ]
        }),
        new TableRow({
          children: [
            this._labelCell('Existing wording'),
            this._valueCell(existing || '[XXXXX]', 6)
          ]
        }),
        new TableRow({
          children: [
            this._labelCell('Proposed revision(s) to\nexisting wording'),
            this._valueCell(proposed || '[XXXXX]', 6)
          ]
        }),
        new TableRow({
          children: [
            this._labelCell(
              'Recommendation(s) with\nrespect to other Works\nTenders/works contracts'
            ),
            this._valueCell(recommendation || '[XXXXX]', 6)
          ]
        })
      ]
    });

    result.push(table);
    return result;
  }

  // ================== עזרות ==================

  private _headerCell(text: string): TableCell {
    return new TableCell({
      children: [
        new Paragraph({
          alignment: AlignmentType.LEFT,
          bidirectional: false,
          children: [new TextRun({ text, bold: true })]
        })
      ]
    });
  }

  private _valueCell(text: string, colSpan: number = 1): TableCell {
    return new TableCell({
      columnSpan: colSpan,
      children: [
        new Paragraph({
          alignment: AlignmentType.LEFT,
          bidirectional: false,
          text: text || ''
        })
      ]
    });
  }

  private _labelCell(text: string): TableCell {
    return new TableCell({
      children: [
        new Paragraph({
          alignment: AlignmentType.LEFT,
          bidirectional: false,
          children: [new TextRun({ text, bold: true })]
        })
      ]
    });
  }

  private async _loadImage(url: string): Promise<ArrayBuffer | undefined> {
    try {
      const response = await fetch(url);
      if (!response.ok) {
        console.warn('Failed to load image (HTTP)', url, response.status);
        return undefined;
      }
      const arrayBuffer = await response.arrayBuffer();
      return arrayBuffer;
    } catch (e) {
      console.warn('Failed to load image', url, e);
      return undefined;
    }
  }

  private _getTextFromRow(row: RowAccessor, internalName: string): string {
    try {
      const raw: any = row.getValueByName(internalName);
      if (raw === null || raw === undefined) return '';

      if (typeof raw === 'boolean') {
        return raw ? 'Yes' : 'No';
      }

      if (typeof raw === 'object') {
        if ((raw as any).Title) return (raw as any).Title;
        if ((raw as any).Label) return (raw as any).Label;
        return this._stripHtml(JSON.stringify(raw));
      }

      if (typeof raw === 'string') {
        const cleanedChoice = this._normalizeChoiceString(raw);
        return this._stripHtml(cleanedChoice);
      }

      return this._stripHtml(String(raw));
    } catch (e) {
      console.warn(`Failed to read field ${internalName} from PMO row`, e);
      return '';
    }
  }

  private _getTextFromIntegration(item: any | null | undefined, internalName: string): string {
    if (!item) return '';
    console.log("🤩 integration item ", item); 
    try {
      const raw: any = item[internalName];
      if (raw === null || raw === undefined) return '';

      if (typeof raw === 'boolean') {
        return raw ? 'Yes' : 'No';
      }

      if (typeof raw === 'string') {
        const cleanedChoice = this._normalizeChoiceString(raw);
        return this._stripHtml(cleanedChoice);
      }

      return this._stripHtml(String(raw));
    } catch (e) {
      console.warn(`Failed to read field ${internalName} from Integration item`, e);
      return '';
    }
  }

  private _normalizeChoiceString(value: string): string {
    const parts = value.split(';#');
    if (parts.length <= 1) return value;

    const meaningful: string[] = [];
    for (let i = 0; i < parts.length; i++) {
      if (i % 2 === 1) {
        meaningful.push(parts[i]);
      }
    }
    return meaningful.join(', ');
  }

  private _stripHtml(html: string): string {
    if (!html) return '';

    let text = html
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/p>\s*<p>/gi, '\n');

    text = text.replace(/<[^>]+>/g, '');

    text = text
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&')
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>')
      .replace(/&quot;/gi, '"')
      .replace(/&#39;/gi, '\'');

    return text.replace(/\s+\n/g, '\n').trim();
  }
}
