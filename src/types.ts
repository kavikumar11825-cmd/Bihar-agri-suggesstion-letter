export interface Suggestion {
  id: number;
  title: string;
  originalText: string;
  detailedContext: string;
  iconName: string;
  stats?: string;
  tag: string;
}

export interface ThemeConfig {
  id: 'green' | 'golden';
  name: string;
  bgImage: string;
  primaryClass: string;
  accentClass: string;
  borderClass: string;
  textClass: string;
  bgMutedClass: string;
  badgeClass: string;
}

export interface BotanicalInfo {
  hindiName: string;
  englishName: string;
  scientificName: string;
  family: string;
  biharStats: string;
  description: string;
  harvestSeason: string;
  iconName: string;
}
