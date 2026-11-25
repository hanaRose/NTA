export interface IOnboardingTip {
  id: number;
  title: string;
  text: string;
  day: number;
  category: string;
  link?: string;
  linkLabel?: string;
  showQr: boolean;
  qrImageUrl?: string;
}
