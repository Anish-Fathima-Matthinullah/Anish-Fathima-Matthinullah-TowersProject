export interface Api {
  id: number;
  identifier: string;
  activityId: string;
  activityName: string;
  blProjectStart: Date | null;
  blProjectFinish: Date | null;
  actualStart: Date | null;
  actualFinish: Date | null;
  activityComplete: number | null;
  materialCostComplete: number | null;
  laborCostComplete: number | null;
  nonLaborCostComplete: number | null;
  importHistoryId: number;
  parentId: string | null;
  color: string;
  font: number;
  bold: boolean;
}
