# Supabase Integration for Revit Family Management

## Overview

This integration enables syncing Revit family data to Supabase, providing:
- **Cloud storage** for family files
- **Database tracking** of family usage and metadata
- **Load history** for analytics and reporting
- **Cross-project** family library management

## Features

### 1. Family Data Syncing
- Automatically sync family metadata to Supabase database
- Track family name, category, file size, and hash
- Record which projects use which families

### 2. File Storage
- Upload family files (.rfa) to Supabase Storage
- Organize files by category in storage buckets
- Support for public or private access

### 3. Load History
- Track when families are loaded into projects
- Record user, project, and timestamp information
- Monitor family usage across teams

## Setup Instructions

### Step 1: Create Supabase Project

1. Go to [supabase.com](https://supabase.com) and create a free account
2. Create a new project
3. Note your project URL (e.g., `https://xxxxx.supabase.co`)

### Step 2: Set Up Database

1. In Supabase dashboard, go to **SQL Editor**
2. Run the schema creation scripts from `SUPABASE_SCHEMA.md`
3. This creates two tables:
   - `families` - stores family metadata
   - `family_load_history` - tracks loading events

### Step 3: Configure Storage

1. In Supabase dashboard, go to **Storage**
2. Create a bucket named `revit-family-free` (or your preferred name)
3. Set bucket to **Public** if you want shareable links, or **Private** for restricted access
4. Configure any storage policies as needed

### Step 4: Get API Credentials

1. In Supabase dashboard, go to **Settings > API**
2. Copy your **Project URL**
3. Copy your **anon/public** API key
4. Keep these secure!

### Step 5: Configure in Revit

1. Open Revit and load the T3Lab Lite extension
2. Click **Supabase Config** button in the Project panel
3. Enter your Project URL and API key
4. Enter your bucket name (default: `revit-family-free`)
5. Click **Test Connection** to verify
6. Click **Save** to store configuration

## Usage

### Loading Families with Sync

1. Click **Load Family** button
2. Select folder containing .rfa files
3. Check the **Sync to Supabase** checkbox
4. Select families to load
5. Click **Load**

The extension will:
- Load families into your Revit project
- Upload family files to Supabase storage (organized by category)
- Create/update database records with metadata
- Record load history with project and user info

### Configuration Management

The extension stores configuration locally in:
```
~/.t3lab/supabase_config.json
```

This file contains:
- Supabase URL
- API key (stored locally only)
- Bucket name

**Important:** Keep your API key secure. Do not commit this file to version control.

## Architecture

### Client Components

1. **supabase_client.py**
   - REST API client for Supabase
   - Uses .NET WebClient (compatible with IronPython)
   - Methods for database and storage operations
   - Singleton pattern for efficient reuse

2. **SupabaseConfigDialog.py**
   - WPF dialog for configuration
   - Test connection functionality
   - Save/load credentials

3. **FamilyLoaderDialog.py**
   - Integrated sync checkbox and config button
   - Automatic syncing after successful loads
   - Error handling and user feedback

### API Methods

#### Database Operations
- `insert_family()` - Add new family record
- `get_family_by_name()` - Find family by name
- `get_family_by_hash()` - Find family by file hash (deduplication)
- `update_family()` - Update existing family
- `search_families()` - Search by name or category
- `insert_load_history()` - Record load event
- `get_load_history()` - Query load history

#### Storage Operations
- `upload_family_file()` - Upload .rfa to storage
- `download_family_file()` - Download .rfa from storage
- `list_storage_files()` - List files in bucket
- `get_public_url()` - Get shareable URL

#### Utility Methods
- `sync_family()` - High-level sync (database + storage)
- `_calculate_file_hash()` - SHA256 hash for deduplication

## Database Schema

### families table
```sql
- id (UUID, primary key)
- name (TEXT)
- file_path (TEXT)
- category (TEXT)
- file_size (BIGINT)
- file_hash (TEXT, indexed)
- created_at (TIMESTAMP)
- updated_at (TIMESTAMP)
- last_loaded_at (TIMESTAMP)
- load_count (INTEGER)
- project_name (TEXT)
- user_name (TEXT)
- metadata (JSONB)
```

### family_load_history table
```sql
- id (UUID, primary key)
- family_id (UUID, foreign key)
- project_name (TEXT)
- user_name (TEXT)
- loaded_at (TIMESTAMP)
- success (BOOLEAN)
- error_message (TEXT)
- revit_version (TEXT)
- metadata (JSONB)
```

## Security Considerations

### Row Level Security (RLS)

The schema includes RLS policies:
- Authenticated users can read all families
- Authenticated users can insert/update their own families
- Load history is readable by authenticated users

### API Key Security

- The anon/public key is used for client-side operations
- This key should only grant access to intended tables
- Use RLS policies to restrict unauthorized access
- Never commit API keys to version control

### Storage Policies

Configure storage bucket policies:
- Public buckets: Anyone can read files
- Private buckets: Only authenticated users can access
- Custom policies for fine-grained control

## Troubleshooting

### Connection Test Fails

1. **Check URL format**: Should be `https://xxxxx.supabase.co` (no trailing slash)
2. **Verify API key**: Copy/paste from Supabase dashboard, no extra spaces
3. **Check network**: Ensure you can reach supabase.co
4. **Review policies**: Make sure RLS policies allow your operations

### Upload Fails

1. **Check bucket name**: Must match exactly (case-sensitive)
2. **Verify storage policies**: Bucket must allow uploads
3. **File size limits**: Check Supabase storage limits for your plan
4. **Network issues**: Large files may timeout on slow connections

### Database Operations Fail

1. **Check schema**: Run the SQL setup scripts in SUPABASE_SCHEMA.md
2. **Verify tables**: Tables `families` and `family_load_history` must exist
3. **Review RLS**: Policies must allow your operations
4. **Check API key**: Must have access to the database

### Sync Checkbox Not Appearing

1. **Check XAML**: Ensure FamilyLoader.xaml includes the checkbox
2. **Restart Revit**: Changes to XAML require reload
3. **Check imports**: Supabase modules must import successfully

## Performance Considerations

### File Upload Speed

- Large family files may take time to upload
- Upload happens after successful load (non-blocking)
- Failed uploads don't prevent family loading
- Consider file size limits for your Supabase plan

### Database Performance

- File hash is indexed for fast deduplication
- Queries use indexes on name, category, hash
- Limit queries to reasonable result sets
- Consider pagination for large datasets

### Caching

- Supabase client uses singleton pattern
- Configuration loaded once per session
- No local caching of family data

## Future Enhancements

Potential improvements:
- Batch upload for better performance
- Download families from Supabase storage
- Browse and search Supabase library
- Team collaboration features
- Usage analytics dashboard
- Automatic conflict resolution
- Thumbnail generation and upload

## Support

For issues or questions:
1. Check this documentation
2. Review Supabase logs in dashboard
3. Check Revit pyRevit logs
4. Open GitHub issue with details

## License

This integration is part of T3Lab Lite extension.

## Credits

- **Author**: T3Lab
- **Supabase**: https://supabase.com
- **pyRevit**: https://github.com/eirannejad/pyRevit
