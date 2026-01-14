export type UserRole = "ADMIN" | "MANAGER" | "BIDDER" | "OBSERVER" | "NONE";

export interface User {
  id: string;
  email: string;
  userName: string;
  role: UserRole;
  name: string;
  avatarUrl?: string | null;
  bio?: string | null;
  isActive: boolean;
  password?: string;
}

export type CalendarProvider = "MICROSOFT" | "GOOGLE";

export interface UserOAuthAccount {
  id: string;
  userId: string;
  provider: string;
  providerAccountId: string;
  email: string;
  displayName?: string | null;
  accessToken?: string | null;
  refreshToken?: string | null;
  expiresAt?: number | null;
  extExpiresIn?: number | null;
  tokenType?: string | null;
  scope?: string | null;
  idToken?: string | null;
  createdAt: string;
  updatedAt: string;
}

export interface ProfileAccount {
  id: string;
  profileId: string;
  provider: CalendarProvider;
  email: string;
  displayName?: string | null;
  timezone?: string | null;
  status?: "ACTIVE" | "INACTIVE";
  lastSyncAt?: string | null;
  createdAt: string;
  updatedAt: string;
}

export type ProfileAccountWithProfile = ProfileAccount & {
  profileDisplayName?: string | null;
  profileAssignedBidderId?: string | null;
};

export interface CalendarEvent {
  id: string;
  title: string;
  start: string;
  end: string;
  isAllDay?: boolean;
  organizer?: string;
  location?: string;
}

export interface BaseInfo extends Record<string, unknown> {
  name?: { first?: string; last?: string };
  contact?: {
    email?: string;
    phone?: string;
    phoneCode?: string;
    phoneNumber?: string;
  };
  links?: Record<string, string> & { linkedin?: string };
  location?: {
    address?: string;
    city?: string;
    state?: string;
    country?: string;
    postalCode?: string;
  };
  career?: {
    jobTitle?: string;
    currentCompany?: string;
    yearsExp?: number | string;
    desiredSalary?: string;
  };
  education?: {
    school?: string;
    degree?: string;
    majorField?: string;
    graduationAt?: string;
  };
  workAuth?: { authorized?: boolean; needsSponsorship?: boolean };
  preferences?: Record<string, unknown>;
  defaultAnswers?: Record<string, string>;
}

export interface LabelAlias {
  id: string;
  canonicalKey: string;
  alias: string;
  normalizedAlias: string;
  isBuiltin?: boolean;
  createdAt?: string;
  updatedAt?: string;
}

export interface Profile {
  id: string;
  displayName: string;
  baseInfo: BaseInfo;
  baseResume?: Record<string, unknown>;
  baseAdditionalBullets?: Record<string, number>;
  createdBy: string;
  createdAt: string;
  updatedAt: string;
  assignedBidderId?: string | null;
  assignedBy?: string | null;
  assignedAt?: string | null;
}

export interface Assignment {
  id: string;
  profileId: string;
  bidderUserId: string;
  assignedBy: string;
  assignedAt: string;
  unassignedAt?: string | null;
}

export type TaskStatus = "todo" | "in_progress" | "in_review" | "done";
export type TaskPriority = "low" | "medium" | "high" | "urgent";
export type TaskAssignmentRequestStatus = "pending" | "approved" | "rejected";
export type TaskDoneRequestStatus = "pending" | "approved" | "rejected";

export interface TaskAssignee {
  id: string;
  name: string;
  email?: string | null;
  avatarUrl?: string | null;
}

export interface Task {
  id: string;
  title: string;
  detail?: string | null;
  status: TaskStatus;
  priority: TaskPriority;
  dueDate: string | null;
  project?: string | null;
  notes?: string | null;
  tags: string[];
  createdBy?: string | null;
  rejectedBy?: string | null;
  rejectedAt?: string | null;
  rejectionReason?: string | null;
  createdAt: string;
  updatedAt: string;
  assignees: TaskAssignee[];
}

export interface TaskAssignmentRequest {
  id: string;
  taskId: string;
  userId: string;
  requestedBy: string;
  status: TaskAssignmentRequestStatus;
  reviewedBy?: string | null;
  reviewedAt?: string | null;
  rejectionReason?: string | null;
  createdAt: string;
  updatedAt: string;
}

export interface TaskAssignmentRequestDetail extends TaskAssignmentRequest {
  taskTitle: string;
  assigneeName?: string | null;
  assigneeEmail?: string | null;
  requesterName?: string | null;
  requesterEmail?: string | null;
}

export interface TaskDoneRequest {
  id: string;
  taskId: string;
  requestedBy: string;
  status: TaskDoneRequestStatus;
  reviewedBy?: string | null;
  reviewedAt?: string | null;
  rejectionReason?: string | null;
  createdAt: string;
  updatedAt: string;
}

export interface TaskDoneRequestDetail extends TaskDoneRequest {
  taskTitle: string;
  requesterName?: string | null;
  requesterEmail?: string | null;
}

export interface ResumeTemplate {
  id: string;
  name: string;
  description?: string | null;
  html: string;
  createdBy?: string | null;
  createdByName?: string | null;
  createdAt: string;
  updatedAt: string;
}

export type SessionStatus =
  | 'OPEN'
  | 'ANALYZED'
  | 'FILLED'
  | 'SUBMITTED'
  | 'ABANDONED'
  | 'ERROR';

export interface ApplicationSession {
  id: string;
  bidderUserId: string;
  profileId: string;
  url: string;
  domain?: string;
  status: SessionStatus;
  jobContext?: Record<string, unknown>;
  formSchema?: Record<string, unknown>;
  fillPlan?: Record<string, unknown>;
  startedAt: string;
  endedAt?: string;
}

export interface ApplicationEvent {
  id: string;
  sessionId: string;
  eventType: string;
  payload?: Record<string, unknown>;
  createdAt: string;
}

export interface ApplicationRecord {
  id: string;
  sessionId: string;
  bidderUserId: string;
  profileId: string;
  resumeId?: string | null;
  url: string;
  domain?: string | null;
  createdAt: string;
}

export interface ApplicationSummary {
  id: string;
  sessionId: string;
  bidderUserId?: string | null;
  bidderName?: string | null;
  bidderEmail?: string | null;
  profileId?: string | null;
  profileDisplayName?: string | null;
  resumeId?: string | null;
  resumeLabel?: string | null;
  url?: string | null;
  domain?: string | null;
  createdAt: string;
}

export type LlmProvider = "OPENAI" | "HUGGINGFACE" | "GEMINI";

export interface LlmSettings {
  id: string;
  ownerType: "ORG" | "USER";
  ownerId: string;
  provider: LlmProvider;
  encryptedApiKey: string;
  chatModel: string;
  embedModel: string;
  updatedAt: string;
}

export type DailyReportStatus = "draft" | "in_review" | "accepted" | "rejected";

export interface DailyReport {
  id: string;
  userId: string;
  reportDate: string;
  status: DailyReportStatus;
  content?: string | null;
  reviewReason?: string | null;
  createdAt: string;
  updatedAt: string;
  submittedAt?: string | null;
  reviewedAt?: string | null;
  reviewedBy?: string | null;
}

export interface DailyReportWithUser extends DailyReport {
  userName: string;
  userEmail: string;
  userAvatarUrl?: string | null;
  userRole?: UserRole;
}

export interface DailyReportAttachment {
  id: string;
  reportId: string;
  fileUrl: string;
  fileName: string;
  fileSize: number;
  mimeType: string;
  createdAt: string;
}

export interface Notification {
  id: string;
  userId: string;
  kind: string;
  message: string;
  href?: string | null;
  createdAt: string;
  readAt?: string | null;
}

export type CommunityThreadType = "CHANNEL" | "DM";

export interface CommunityThread {
  id: string;
  threadType: CommunityThreadType;
  name?: string | null;
  nameKey?: string | null;
  description?: string | null;
  createdBy?: string | null;
  isPrivate: boolean;
  createdAt: string;
  lastMessageAt?: string | null;
}


export interface CommunityThreadParticipant {
  id: string;
  name: string;
  email: string;
  avatarUrl?: string | null;
}

export interface CommunityThreadSummary {
  id: string;
  threadType: CommunityThreadType;
  name?: string | null;
  description?: string | null;
  isPrivate: boolean;
  createdAt: string;
  lastMessageAt?: string | null;
  participants?: CommunityThreadParticipant[];
}

export interface CommunityMessage {
  id: string;
  threadId: string;
  senderId: string;
  senderName?: string | null;
  senderAvatarUrl?: string | null;
  body: string;
  replyToMessageId?: string | null;
  isEdited: boolean;
  editedAt?: string | null;
  isDeleted: boolean;
  deletedAt?: string | null;
  createdAt: string;
}

export interface MessageAttachment {
  id: string;
  messageId: string;
  fileName: string;
  fileUrl: string;
  fileSize: number;
  mimeType: string;
  thumbnailUrl?: string | null;
  width?: number | null;
  height?: number | null;
  createdAt: string;
}

export interface MessageReaction {
  id: string;
  messageId: string;
  userId: string;
  emoji: string;
  createdAt: string;
}

export interface ReactionSummary {
  emoji: string;
  count: number;
  userIds: string[];
  hasCurrentUser: boolean;
}

export interface UnreadInfo {
  threadId: string;
  unreadCount: number;
  lastReadMessageId?: string | null;
  lastReadAt: string;
}

export interface UserPresence {
  userId: string;
  status: "online" | "away" | "busy" | "offline";
  lastSeenAt: string;
}

export interface PinnedMessage {
  id: string;
  threadId: string;
  messageId: string;
  pinnedBy: string;
  pinnedAt: string;
}

export interface ChannelPermissions {
  can_post?: boolean;
  can_invite?: boolean;
  can_delete_messages?: boolean;
  can_pin_messages?: boolean;
  can_manage_roles?: boolean;
}

export interface ChannelRole {
  id: string;
  channelId: string;
  roleName: string;
  permissions: ChannelPermissions;
  createdAt: string;
}

export interface ThreadMemberWithPermissions {
  id: string;
  threadId: string;
  userId: string;
  role: string;
  permissions: ChannelPermissions;
  joinedAt: string;
}

export interface MessageReadReceipt {
  id: string;
  messageId: string;
  userId: string;
  readAt: string;
}

export interface ReadReceiptSummary {
  totalReaders: number;
  userIds: string[];
}

export interface CommunityMessageExtended extends CommunityMessage {
  attachments?: MessageAttachment[];
  reactions?: ReactionSummary[];
  replyPreview?: {
    id: string;
    senderId: string;
    senderName?: string | null;
    body: string;
  } | null;
  readReceipts?: ReadReceiptSummary;
}
