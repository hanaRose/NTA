declare interface IShipWorkCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ShipWorkCommandSetStrings' {
  const strings: IShipWorkCommandSetStrings;
  export = strings;
}
