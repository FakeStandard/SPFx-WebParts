import * as React from 'react';
import styles from './IconSample.module.scss';
import { IIconSampleProps } from './IIconSampleProps';
import { IDocument, IIconSampleStates } from './IconSampleStates'
import { escape } from '@microsoft/sp-lodash-subset';
import { DetailsList, IColumn, mergeStyleSets, SelectionMode, Toggle, TooltipHost } from '@fluentui/react';

const classNames = mergeStyleSets({
  fileIconHeaderIcon: {
    padding: 0,
    fontSize: '16px',
  },
  fileIconCell: {
    textAlign: 'center',
    selectors: {
      '&:before': {
        content: '.',
        display: 'inline-block',
        verticalAlign: 'middle',
        height: '100%',
        width: '0px',
        visibility: 'hidden',
      },
    },
  },
  fileIconImg: {
    verticalAlign: 'middle',
    maxHeight: '16px',
    maxWidth: '16px',
  },
  controlWrapper: {
    display: 'flex',
    flexWrap: 'wrap',
  },
  exampleToggle: {
    display: 'inline-block',
    marginBottom: '10px',
    marginRight: '30px',
  },
  selectionDetails: {
    marginBottom: '20px',
  },
});

export default class IconSample extends React.Component<IIconSampleProps, IIconSampleStates> {

  private items: IDocument[];
  private columns: IColumn[];
  constructor(props: IIconSampleProps) {
    super(props);

    this.items = _generateDocuments();
    this.columns = [
      {
        key: '1',
        name: 'File Type',
        iconClassName: '',
        ariaLabel: '',
        iconName: 'Page',
        isIconOnly: true,
        fieldName: 'name',
        minWidth: 16,
        maxWidth: 16,
        onRender: (item: IDocument) => (
          <TooltipHost content={`${item.fileType} file`}>
            <img src={item.iconName} className={classNames.fileIconImg} alt={`${item.fileType} file icon`} />
          </TooltipHost>
        ),
      },
      {
        key: '2',
        name: 'Name',
        fieldName: 'name',
        minWidth: 400,
        maxWidth: 400,
        isRowHeader: true,
        isResizable: false,
        data: 'string',
        isPadded: true,
      },
      {
        key: '3',
        name: 'Date Modified',
        fieldName: 'dateModifiedValue',
        minWidth: 100,
        maxWidth: 100,
        isResizable: false,
        data: 'number',
        onRender: (item: IDocument) => {
          return <span>{item.dateModified}</span>;
        },
        isPadded: true,
      },
      {
        key: '4',
        name: 'Modified By',
        fieldName: 'modifiedBy',
        minWidth: 130,
        maxWidth: 130,
        isResizable: false,
        isCollapsible: true,
        data: 'string',
        onRender: (item: IDocument) => {
          return <span>{item.modifiedBy}</span>;
        },
        isPadded: true,
      },
      {
        key: '5',
        name: 'File Size',
        fieldName: 'fileSizeRaw',
        minWidth: 90,
        maxWidth: 90,
        isResizable: false,
        isCollapsible: true,
        data: 'number',
        onRender: (item: IDocument) => {
          return <span>{item.fileSize}</span>;
        },
      },
    ]

    this.state = {
      columns: this.columns,
      items: this.items,
      isCompactMode: false,
    }
  }

  public render(): React.ReactElement<IIconSampleProps> {
    const { isCompactMode } = this.state;

    return (
      <div className={styles.iconSample}>
        <Toggle label="Enable compact mode" checked={isCompactMode} onText="Compact" offText="Normal"
          onChange={(e, checked) => this.setState({ isCompactMode: checked })} />
        <DetailsList
          columns={this.columns}
          items={this.items}
          compact={isCompactMode}
          selectionMode={SelectionMode.none} />
      </div>
    );
  }
}

function _generateDocuments() {
  const items: IDocument[] = [];
  for (let i = 0; i < 500; i++) {
    const randomDate = _randomDate(new Date(2012, 0, 1), new Date());
    const randomFileSize = _randomFileSize();
    const randomFileType = _randomFileIcon();
    let fileName = _lorem(2);
    fileName = fileName.charAt(0).toUpperCase() + fileName.slice(1).concat(`.${randomFileType.docType}`);
    let userName = _lorem(2);
    userName = userName
      .split(' ')
      .map((name: string) => name.charAt(0).toUpperCase() + name.slice(1))
      .join(' ');
    items.push({
      key: i.toString(),
      name: fileName,
      value: fileName,
      iconName: randomFileType.url,
      fileType: randomFileType.docType,
      modifiedBy: userName,
      dateModified: randomDate.dateFormatted,
      dateModifiedValue: randomDate.value,
      fileSize: randomFileSize.value,
      fileSizeRaw: randomFileSize.rawSize,
    });
  }
  return items;
}

function _randomDate(start: Date, end: Date): { value: number; dateFormatted: string } {
  const date: Date = new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
  return {
    value: date.valueOf(),
    dateFormatted: date.toLocaleDateString(),
  };
}

function _randomFileSize(): { value: string; rawSize: number } {
  const fileSize: number = Math.floor(Math.random() * 100) + 30;
  return {
    value: `${fileSize} KB`,
    rawSize: fileSize,
  };
}

function _randomFileIcon(): { docType: string; url: string } {
  const docType: string = FILE_ICONS[Math.floor(Math.random() * FILE_ICONS.length)].name;
  return {
    docType,
    url: `https://static2.sharepointonline.com/files/fabric/assets/item-types/16/${docType}.svg`,
  };
}

function _lorem(wordCount: number): string {
  const startIndex = loremIndex + wordCount > LOREM_IPSUM.length ? 0 : loremIndex;
  loremIndex = startIndex + wordCount;
  return LOREM_IPSUM.slice(startIndex, loremIndex).join(' ');
}

const FILE_ICONS: { name: string }[] = [
  { name: 'accdb' },
  { name: 'audio' },
  { name: 'code' },
  { name: 'csv' },
  { name: 'docx' },
  { name: 'dotx' },
  { name: 'mpp' },
  { name: 'mpt' },
  { name: 'model' },
  { name: 'one' },
  { name: 'onetoc' },
  { name: 'potx' },
  { name: 'ppsx' },
  { name: 'pdf' },
  { name: 'photo' },
  { name: 'pptx' },
  { name: 'presentation' },
  { name: 'potx' },
  { name: 'pub' },
  { name: 'rtf' },
  { name: 'spreadsheet' },
  { name: 'txt' },
  { name: 'vector' },
  { name: 'vsdx' },
  { name: 'vssx' },
  { name: 'vstx' },
  { name: 'xlsx' },
  { name: 'xltx' },
  { name: 'xsn' },
];

let loremIndex = 0;
const LOREM_IPSUM = (
  'lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut ' +
  'labore et dolore magna aliqua ut enim ad minim veniam quis nostrud exercitation ullamco laboris nisi ut ' +
  'aliquip ex ea commodo consequat duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore ' +
  'eu fugiat nulla pariatur excepteur sint occaecat cupidatat non proident sunt in culpa qui officia deserunt '
).split(' ');