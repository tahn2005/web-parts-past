declare interface ILinkCardsWebPartStrings {
  // On the page
  OpensInNewTab: string;
  TextPlaceholder: string;
  UntitledCard: string;
  NoLinkYet: string;

  // Edit-mode canvas controls
  AddCard: string;
  EditCard: string;
  MoveCardEarlier: string;
  MoveCardLater: string;
  DeleteCard: string;
  CardOptions: string;
  ConfirmDeleteCard: string;

  // Property pane: section
  SectionPaneDescription: string;
  CardsGroup: string;
  CardsHelp: string;
  EditCardButton: string;
  AppearanceGroup: string;
  PaletteField: string;
  PaletteBlue: string;
  PaletteIndigo: string;
  PaletteGreen: string;
  PaletteAmber: string;
  PaletteCustom: string;
  IconColorField: string;
  TextColorField: string;
  CardBackgroundField: string;
  ColumnsField: string;
  ColumnsAuto: string;
  ColumnsOne: string;
  ColumnsTwo: string;
  ColumnsThree: string;
  FontGroup: string;
  UseInterFontField: string;
  AccessibilityGroup: string;
  AccessibilityAllPass: string;
  AccessibilityAdjusted: string;

  // Property pane: card
  CardPaneDescription: string;
  BackToSection: string;
  CardTextGroup: string;
  CardTextField: string;
  CardLinkGroup: string;
  LinkUrlField: string;
  LinkUrlHelp: string;
  OpenInNewTabField: string;
  CardIconGroup: string;
  IconField: string;
  IconColorModeField: string;
  IconColorSection: string;
  IconColorCustom: string;
  CardIconColorField: string;
  ArrangeGroup: string;
  DuplicateCard: string;
}

declare module 'LinkCardsWebPartStrings' {
  const strings: ILinkCardsWebPartStrings;
  export = strings;
}
