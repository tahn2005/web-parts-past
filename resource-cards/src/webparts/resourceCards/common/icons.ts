/**
 * Icons authors can choose from.
 *
 * The same Lucide icons the original React app uses, imported one by one
 * so only these end up in the bundle. The first block is every icon used on
 * the four lifecycle pages; the second adds general-purpose ones.
 *
 * To offer another icon: import it from 'lucide-react' and add a row below.
 * The `name` is what gets saved with the page, so never rename an existing one.
 */

import {
  AlertCircle, AlertTriangle, ArrowRightLeft, Award, Banknote, BarChart2, BarChart3, Book, BookOpen,
  Briefcase, Calendar, CalendarCheck, CalendarClock, CalendarDays, ClipboardCheck, ClipboardList, Coffee,
  Compass, CreditCard, DollarSign, FileSignature, FileText, FileWarning, Gavel, Gift, Globe, GraduationCap,
  Handshake, HardHat, Heart, HeartHandshake, HeartPulse, KeyRound, Layers, Library, LifeBuoy, Lightbulb,
  ListChecks, LogOut, Map, Megaphone, MessageSquare, MessagesSquare, Mic2, Monitor, Network, Newspaper,
  PartyPopper, PiggyBank, Plane, Presentation, Receipt, RefreshCw, Rocket, Route, Scale, Search, Settings,
  ShieldCheck, Smile, Sparkles, Speech, Star, Target, TrendingUp, Umbrella, UserCheck, UserPlus, Users,
  Video, Workflow,
  Bell, Bookmark, Building2, CheckCircle, Clock, Contact, ExternalLink, Flag, Folder, HelpCircle, Home,
  Info, Laptop, Link, Mail, Phone, ThumbsUp, Trophy, UserCog, UserMinus, Wrench
} from 'lucide-react';
import type { LucideIcon } from 'lucide-react';

export interface IIconOption {
  name: string;
  label: string;
  /** Extra search terms. */
  keywords: string;
  Icon: LucideIcon;
}

export const ICONS: IIconOption[] = [
  // Used on the original lifecycle pages
  { name: 'AlertCircle', label: 'Alert circle', keywords: 'warning notice attention', Icon: AlertCircle },
  { name: 'AlertTriangle', label: 'Warning', keywords: 'alert caution risk', Icon: AlertTriangle },
  { name: 'ArrowRightLeft', label: 'Transfer', keywords: 'swap move internal mobility', Icon: ArrowRightLeft },
  { name: 'Award', label: 'Award', keywords: 'recognition badge prize', Icon: Award },
  { name: 'Banknote', label: 'Banknote', keywords: 'money pay cash salary', Icon: Banknote },
  { name: 'BarChart2', label: 'Bar chart', keywords: 'report analytics data', Icon: BarChart2 },
  { name: 'BarChart3', label: 'Bar chart (alt)', keywords: 'report analytics planning data', Icon: BarChart3 },
  { name: 'Book', label: 'Book', keywords: 'handbook manual policy', Icon: Book },
  { name: 'BookOpen', label: 'Open book', keywords: 'learning guide read', Icon: BookOpen },
  { name: 'Briefcase', label: 'Briefcase', keywords: 'job work role', Icon: Briefcase },
  { name: 'Calendar', label: 'Calendar', keywords: 'date event schedule', Icon: Calendar },
  { name: 'CalendarCheck', label: 'Calendar check', keywords: 'check-in milestone schedule', Icon: CalendarCheck },
  { name: 'CalendarClock', label: 'Calendar clock', keywords: 'meeting 1:1 recurring', Icon: CalendarClock },
  { name: 'CalendarDays', label: 'Calendar days', keywords: 'schedule leave time off', Icon: CalendarDays },
  { name: 'ClipboardCheck', label: 'Clipboard check', keywords: 'feedback review complete', Icon: ClipboardCheck },
  { name: 'ClipboardList', label: 'Clipboard list', keywords: 'checklist onboarding tasks', Icon: ClipboardList },
  { name: 'Coffee', label: 'Coffee', keywords: 'break rest meal', Icon: Coffee },
  { name: 'Compass', label: 'Compass', keywords: 'direction competencies guide', Icon: Compass },
  { name: 'CreditCard', label: 'Credit card', keywords: 'corporate card expense payment', Icon: CreditCard },
  { name: 'DollarSign', label: 'Dollar sign', keywords: 'money compensation bonus pay', Icon: DollarSign },
  { name: 'FileSignature', label: 'Signed document', keywords: 'offer contract agreement', Icon: FileSignature },
  { name: 'FileText', label: 'Document', keywords: 'file page description', Icon: FileText },
  { name: 'FileWarning', label: 'Document warning', keywords: 'corrective action notice', Icon: FileWarning },
  { name: 'Gavel', label: 'Gavel', keywords: 'legal law compliance', Icon: Gavel },
  { name: 'Gift', label: 'Gift', keywords: 'benefits referral bonus', Icon: Gift },
  { name: 'Globe', label: 'Globe', keywords: 'world global international', Icon: Globe },
  { name: 'GraduationCap', label: 'Graduation cap', keywords: 'learning education internship training', Icon: GraduationCap },
  { name: 'Handshake', label: 'Handshake', keywords: 'partnership agreement relations', Icon: Handshake },
  { name: 'HardHat', label: 'Hard hat', keywords: 'safety injury workplace', Icon: HardHat },
  { name: 'Heart', label: 'Heart', keywords: 'care wellbeing support', Icon: Heart },
  { name: 'HeartHandshake', label: 'Heart handshake', keywords: 'support care coaching', Icon: HeartHandshake },
  { name: 'HeartPulse', label: 'Heart pulse', keywords: 'health wellbeing medical', Icon: HeartPulse },
  { name: 'KeyRound', label: 'Key', keywords: 'access systems password', Icon: KeyRound },
  { name: 'Layers', label: 'Layers', keywords: 'types classification levels', Icon: Layers },
  { name: 'Library', label: 'Library', keywords: 'resources collection', Icon: Library },
  { name: 'LifeBuoy', label: 'Life buoy', keywords: 'help support assistance', Icon: LifeBuoy },
  { name: 'Lightbulb', label: 'Light bulb', keywords: 'idea tips learning', Icon: Lightbulb },
  { name: 'ListChecks', label: 'Checklist', keywords: 'responsibilities tasks list', Icon: ListChecks },
  { name: 'LogOut', label: 'Exit', keywords: 'leave offboarding separation', Icon: LogOut },
  { name: 'Map', label: 'Map', keywords: 'roadmap plan route', Icon: Map },
  { name: 'Megaphone', label: 'Megaphone', keywords: 'announcement communication news', Icon: Megaphone },
  { name: 'MessageSquare', label: 'Message', keywords: 'chat feedback comment', Icon: MessageSquare },
  { name: 'MessagesSquare', label: 'Conversation', keywords: 'chat interview discussion', Icon: MessagesSquare },
  { name: 'Mic2', label: 'Microphone', keywords: 'speaker podcast talk', Icon: Mic2 },
  { name: 'Monitor', label: 'Monitor', keywords: 'computer system software', Icon: Monitor },
  { name: 'Network', label: 'Network', keywords: 'organization structure hierarchy', Icon: Network },
  { name: 'Newspaper', label: 'Newspaper', keywords: 'news report update', Icon: Newspaper },
  { name: 'PartyPopper', label: 'Party popper', keywords: 'celebration event mixer', Icon: PartyPopper },
  { name: 'PiggyBank', label: 'Piggy bank', keywords: 'savings retirement 401k', Icon: PiggyBank },
  { name: 'Plane', label: 'Plane', keywords: 'travel trip flight', Icon: Plane },
  { name: 'Presentation', label: 'Presentation', keywords: 'meeting slides all staff', Icon: Presentation },
  { name: 'Receipt', label: 'Receipt', keywords: 'expense reimbursement', Icon: Receipt },
  { name: 'RefreshCw', label: 'Refresh', keywords: 'renew cycle change', Icon: RefreshCw },
  { name: 'Rocket', label: 'Rocket', keywords: 'launch growth program', Icon: Rocket },
  { name: 'Route', label: 'Route', keywords: 'path career journey', Icon: Route },
  { name: 'Scale', label: 'Scale', keywords: 'ethics balance fairness legal', Icon: Scale },
  { name: 'Search', label: 'Search', keywords: 'find recruiting look', Icon: Search },
  { name: 'Settings', label: 'Settings', keywords: 'gear configuration tools', Icon: Settings },
  { name: 'ShieldCheck', label: 'Shield check', keywords: 'compliance security policy', Icon: ShieldCheck },
  { name: 'Smile', label: 'Smile', keywords: 'happy engagement wellbeing', Icon: Smile },
  { name: 'Sparkles', label: 'Sparkles', keywords: 'new highlight special', Icon: Sparkles },
  { name: 'Speech', label: 'Speech', keywords: 'coaching talk conversation', Icon: Speech },
  { name: 'Star', label: 'Star', keywords: 'favorite featured recognition', Icon: Star },
  { name: 'Target', label: 'Target', keywords: 'goals objectives aim', Icon: Target },
  { name: 'TrendingUp', label: 'Trending up', keywords: 'growth performance development', Icon: TrendingUp },
  { name: 'Umbrella', label: 'Umbrella', keywords: 'insurance protection leave', Icon: Umbrella },
  { name: 'UserCheck', label: 'Person check', keywords: 'hire approve selection', Icon: UserCheck },
  { name: 'UserPlus', label: 'Add person', keywords: 'hire onboard new backfill', Icon: UserPlus },
  { name: 'Users', label: 'People', keywords: 'team group employees', Icon: Users },
  { name: 'Video', label: 'Video', keywords: 'recording training meeting', Icon: Video },
  { name: 'Workflow', label: 'Workflow', keywords: 'process steps flow', Icon: Workflow },

  // General purpose
  { name: 'Bell', label: 'Bell', keywords: 'notification reminder alert', Icon: Bell },
  { name: 'Bookmark', label: 'Bookmark', keywords: 'save favorite', Icon: Bookmark },
  { name: 'Building2', label: 'Building', keywords: 'office company venue', Icon: Building2 },
  { name: 'CheckCircle', label: 'Check circle', keywords: 'done complete approved', Icon: CheckCircle },
  { name: 'Clock', label: 'Clock', keywords: 'time attendance hours', Icon: Clock },
  { name: 'Contact', label: 'Contact card', keywords: 'person directory profile', Icon: Contact },
  { name: 'ExternalLink', label: 'External link', keywords: 'open website', Icon: ExternalLink },
  { name: 'Flag', label: 'Flag', keywords: 'milestone report', Icon: Flag },
  { name: 'Folder', label: 'Folder', keywords: 'files documents library', Icon: Folder },
  { name: 'HelpCircle', label: 'Help', keywords: 'question faq support', Icon: HelpCircle },
  { name: 'Home', label: 'Home', keywords: 'house start main', Icon: Home },
  { name: 'Info', label: 'Information', keywords: 'about details', Icon: Info },
  { name: 'Laptop', label: 'Laptop', keywords: 'computer remote device', Icon: Laptop },
  { name: 'Link', label: 'Link', keywords: 'url chain', Icon: Link },
  { name: 'Mail', label: 'Mail', keywords: 'email contact message', Icon: Mail },
  { name: 'Phone', label: 'Phone', keywords: 'call contact help line', Icon: Phone },
  { name: 'ThumbsUp', label: 'Thumbs up', keywords: 'like approve feedback', Icon: ThumbsUp },
  { name: 'Trophy', label: 'Trophy', keywords: 'achievement win recognition', Icon: Trophy },
  { name: 'UserCog', label: 'Person settings', keywords: 'manager admin hr', Icon: UserCog },
  { name: 'UserMinus', label: 'Remove person', keywords: 'termination separation exit', Icon: UserMinus },
  { name: 'Wrench', label: 'Wrench', keywords: 'tools fix maintenance', Icon: Wrench }
];

const ICONS_BY_NAME: { [name: string]: IIconOption } = {};
ICONS.forEach((option) => {
  ICONS_BY_NAME[option.name] = option;
});

/** Looks up an icon by its saved name, falling back to the document icon. */
export function getIcon(name: string | undefined): IIconOption {
  return (name && ICONS_BY_NAME[name]) || ICONS_BY_NAME.FileText;
}

export function searchIcons(query: string): IIconOption[] {
  const terms = query.toLowerCase().split(/\s+/).filter(Boolean);
  if (terms.length === 0) {
    return ICONS;
  }
  return ICONS.filter((option) => {
    const haystack = `${option.name} ${option.label} ${option.keywords}`.toLowerCase();
    return terms.every((term) => haystack.indexOf(term) !== -1);
  });
}
