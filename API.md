# Backend API Documentation

This document lists all available API endpoints in the backend.

## Base URL
All endpoints are prefixed with the backend server URL (e.g., `http://localhost:3000`)

## Authentication
Most endpoints require authentication via Bearer token in the `Authorization` header:
```
Authorization: Bearer <token>
```

Public endpoints (no authentication required):
- `/health`
- `/auth/login`
- `/auth/signup`

## API Endpoints

### Health & Status
- `GET /health` - Health check endpoint

### Authentication
- `POST /auth/login` - User login
- `POST /auth/signup` - User registration

### Profiles
- `GET /profiles` - List all profiles (filtered by user role)
- `POST /profiles` - Create a new profile
- `PATCH /profiles/:id` - Update a profile

### Resume Templates
- `GET /resume-templates` - List all resume templates
- `POST /resume-templates` - Create a new resume template
- `PATCH /resume-templates/:id` - Update a resume template
- `DELETE /resume-templates/:id` - Delete a resume template
- `POST /resume-templates/render-pdf` - Render resume template to PDF

### Tasks
- `GET /tasks` - List tasks
- `POST /tasks` - Create a new task
- `PATCH /tasks/:id` - Update a task
- `PATCH /tasks/:id/notes` - Update task notes
- `DELETE /tasks/:id` - Delete a task
- `GET /tasks/requests` - Get task requests
- `POST /tasks/:id/approve` - Approve a task
- `POST /tasks/:id/reject` - Reject a task
- `POST /tasks/:id/assign-requests` - Create assign request for a task
- `GET /tasks/assign-requests` - Get assign requests
- `POST /tasks/:id/done-requests` - Create done request for a task
- `GET /tasks/done-requests` - Get done requests
- `POST /tasks/done-requests/:id/approve` - Approve a done request
- `POST /tasks/done-requests/:id/reject` - Reject a done request
- `POST /tasks/assign-requests/:id/approve` - Approve an assign request
- `POST /tasks/assign-requests/:id/reject` - Reject an assign request
- `POST /tasks/:id/assign-self-request` - Request self-assignment for a task
- `POST /tasks/:id/assign-self` - Assign task to self

### Calendar & Events
- `GET /calendar/accounts` - List calendar accounts
- `POST /calendar/accounts` - Add calendar account
- `POST /calendar/events/sync` - Sync calendar events
- `GET /calendar/events/stored` - Get stored calendar events
- `GET /calendar/events` - Get calendar events

### Daily Reports
- `GET /daily-reports` - List daily reports
- `GET /daily-reports/by-date` - Get daily report by date
- `PUT /daily-reports/by-date` - Create or update daily report by date
- `POST /daily-reports/by-date/send` - Send daily report
- `PATCH /daily-reports/:id/status` - Update daily report status
- `GET /daily-reports/:id/attachments` - Get daily report attachments
- `POST /daily-reports/upload` - Upload attachment for daily report

### Notifications
- `GET /notifications/summary` - Get notification summary
- `GET /notifications/list` - List notifications

### Admin - Daily Reports
- `GET /admin/daily-reports/by-date` - Get daily reports by date (admin)
- `GET /admin/daily-reports/in-review` - Get daily reports in review (admin)
- `GET /admin/daily-reports/accepted-by-date` - Get accepted daily reports by date (admin)
- `GET /admin/daily-reports/by-user` - Get daily reports by user (admin)

### Assignments
- `GET /assignments` - List assignments
- `POST /assignments` - Create an assignment
- `POST /assignments/:id/unassign` - Unassign an assignment

### Community
- `GET /community/overview` - Get community overview
- `GET /community/channels` - List community channels
- `POST /community/channels` - Create a community channel
- `PATCH /community/channels/:id` - Update a community channel
- `DELETE /community/channels/:id` - Delete a community channel
- `POST /community/dms` - Create or get DM thread
- `GET /community/threads/:id/messages` - Get messages in a thread
- `POST /community/threads/:id/messages` - Send message in a thread
- `PATCH /community/messages/:messageId` - Edit a message
- `DELETE /community/messages/:messageId` - Delete a message
- `POST /community/messages/:messageId/reactions` - Add reaction to message
- `DELETE /community/messages/:messageId/reactions` - Remove reaction from message
- `POST /community/messages/:messageId/pin` - Pin a message
- `DELETE /community/messages/:messageId/pin` - Unpin a message
- `GET /community/threads/:id/pins` - Get pinned messages in thread
- `POST /community/threads/:id/mark-read` - Mark thread as read
- `POST /community/messages/mark-read` - Mark messages as read
- `POST /community/upload` - Upload file to community
- `GET /community/unread-summary` - Get unread message summary
- `POST /community/presence` - Update user presence
- `GET /community/presence` - Get user presence

### WebSockets - Community
- `GET /ws/community` - WebSocket connection for community features (real-time chat, typing indicators)

### Application Sessions
- `GET /sessions` - List application sessions
- `GET /sessions/:id` - Get application session by ID
- `POST /sessions` - Create a new application session
- `POST /sessions/:id/go` - Navigate to URL in session
- `POST /sessions/:id/analyze` - Analyze page in session
- `POST /sessions/:id/autofill` - Autofill form in session
- `POST /sessions/:id/mark-submitted` - Mark session as submitted

### WebSockets - Browser
- `GET /ws/browser/:sessionId` - WebSocket connection for live browser session streaming

### LLM Services
- `POST /llm/resume-parse` - Parse resume using LLM
- `POST /llm/job-analyze` - Analyze job posting using LLM
- `POST /llm/rank-resumes` - Rank resumes using LLM
- `POST /llm/autofill-plan` - Generate autofill plan using LLM
- `POST /llm/tailor-resume` - Tailor resume using LLM

### Label Aliases
- `GET /label-aliases` - List label aliases
- `POST /label-aliases` - Create label alias
- `PATCH /label-aliases/:id` - Update label alias
- `DELETE /label-aliases/:id` - Delete label alias
- `GET /application-phrases` - Get application phrases

### Users
- `GET /users` - List users
- `PATCH /users/:id/role` - Update user role
- `POST /users/me/avatar` - Upload user avatar

### Metrics
- `GET /metrics/my` - Get user metrics

### Settings
- `GET /settings/llm` - Get LLM settings
- `POST /settings/llm` - Update LLM settings

### Manager
- `GET /manager/bidders/summary` - Get bidder summaries (manager only)
- `GET /manager/applications` - Get applications (manager only)

### Scraper APIs
- `GET /scraper/job-links` - List job links with filtering
- `GET /scraper/countries` - List countries

## Notes

### Role-Based Access
- **OBSERVER**: Limited read-only access
- **BIDDER**: Can manage assigned profiles and tasks
- **MANAGER**: Can manage bidders and view all applications
- **ADMIN**: Full access to all endpoints

### WebSocket Endpoints
- `/ws/community` - Requires authentication token in query parameter or upgrade header
- `/ws/browser/:sessionId` - Currently allows connections without authentication for demo purposes

### File Uploads
- Daily report attachments: `POST /daily-reports/upload`
- Community files: `POST /community/upload`
- User avatars: `POST /users/me/avatar`

All file uploads use multipart/form-data encoding with a 10MB file size limit.
