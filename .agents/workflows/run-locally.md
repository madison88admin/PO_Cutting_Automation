---
description: How to run the PO Cutting Automation system on localhost
---

To run the system locally, follow these steps:

### 1. Prerequisites
- Ensure you have **Node.js** (v18 or higher) installed.
- Ensure you have a **Supabase** project created.

### 2. Environment Setup
Create or update the `.env.local` file in the root directory with your Supabase credentials.

```env
NEXT_PUBLIC_SUPABASE_URL=your_supabase_project_url
NEXT_PUBLIC_SUPABASE_ANON_KEY=your_supabase_anon_key
SUPABASE_SERVICE_ROLE_KEY=your_supabase_service_role_key
DATA_SOURCE=upload
```

### 3. Database Schema
Ensure the following tables exist in your Supabase project:
- `users` (id, name, email, role, is_active)
- `audit_logs` (id, event, user_id, run_id, metadata, ip_address, created_at)
- `run_history` (id, user_id, filename, status, error_count, warning_count, orders_rows, lines_rows, order_sizes_rows, created_at, completed_at)
- `factory_mapping`, `mlo_mapping`, `column_mapping`

### 4. Installation
Open your terminal in the project root and run:
// turbo
```bash
npm install
```

### 5. Start Development
Run the development server:
// turbo
```bash
npm run dev
```

### 6. Access the App
Open [http://localhost:3000](http://localhost:3000) in your browser.
