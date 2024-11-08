declare interface IXeokitViewerCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'XeokitViewerCommandSetStrings' {
  const strings: IXeokitViewerCommandSetStrings;
  export = strings;
}
