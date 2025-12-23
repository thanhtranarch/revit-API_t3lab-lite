# Supabase Database Schema for Revit Families

## Overview
This document describes the database schema for storing Revit family information in Supabase.

## Tables

### 1. families
Stores information about Revit family files.

```sql
CREATE TABLE families (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    name TEXT NOT NULL,
    file_path TEXT NOT NULL,
    category TEXT,
    file_size BIGINT,
    file_hash TEXT,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    updated_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    last_loaded_at TIMESTAMP WITH TIME ZONE,
    load_count INTEGER DEFAULT 0,
    project_name TEXT,
    user_name TEXT,
    metadata JSONB
);

-- Index for faster searches
CREATE INDEX idx_families_name ON families(name);
CREATE INDEX idx_families_category ON families(category);
CREATE INDEX idx_families_created_at ON families(created_at);
CREATE INDEX idx_families_file_hash ON families(file_hash);
```

### 2. family_load_history
Tracks when families are loaded into Revit projects.

```sql
CREATE TABLE family_load_history (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    family_id UUID REFERENCES families(id) ON DELETE CASCADE,
    project_name TEXT NOT NULL,
    user_name TEXT NOT NULL,
    loaded_at TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
    success BOOLEAN DEFAULT TRUE,
    error_message TEXT,
    revit_version TEXT,
    metadata JSONB
);

-- Index for performance
CREATE INDEX idx_load_history_family_id ON family_load_history(family_id);
CREATE INDEX idx_load_history_loaded_at ON family_load_history(loaded_at);
CREATE INDEX idx_load_history_project ON family_load_history(project_name);
```

## Row Level Security (RLS)

### Enable RLS
```sql
ALTER TABLE families ENABLE ROW LEVEL SECURITY;
ALTER TABLE family_load_history ENABLE ROW LEVEL SECURITY;
```

### Policies
```sql
-- Allow all authenticated users to read families
CREATE POLICY "Allow authenticated read on families"
ON families FOR SELECT
TO authenticated
USING (true);

-- Allow authenticated users to insert their own families
CREATE POLICY "Allow authenticated insert on families"
ON families FOR INSERT
TO authenticated
WITH CHECK (true);

-- Allow authenticated users to update families
CREATE POLICY "Allow authenticated update on families"
ON families FOR UPDATE
TO authenticated
USING (true);

-- Allow all authenticated users to read load history
CREATE POLICY "Allow authenticated read on family_load_history"
ON family_load_history FOR SELECT
TO authenticated
USING (true);

-- Allow authenticated users to insert load history
CREATE POLICY "Allow authenticated insert on family_load_history"
ON family_load_history FOR INSERT
TO authenticated
WITH CHECK (true);
```

## Setup Instructions

1. Create a Supabase project at https://supabase.com
2. Go to SQL Editor and run the schema creation scripts above
3. Get your project URL and anon/public API key from Settings > API
4. Configure the credentials in the Revit extension

## API Endpoints

The Supabase REST API will be used with these endpoints:

- `POST /rest/v1/families` - Insert new family
- `GET /rest/v1/families?name=eq.{name}` - Get family by name
- `PATCH /rest/v1/families?id=eq.{id}` - Update family
- `POST /rest/v1/family_load_history` - Insert load history
- `GET /rest/v1/families?select=*&order=name.asc` - Get all families

## Metadata JSONB Fields

### families.metadata
```json
{
    "thumbnail_url": "string",
    "parameters": ["param1", "param2"],
    "tags": ["tag1", "tag2"],
    "description": "string",
    "revit_version": "2024",
    "file_extension": ".rfa"
}
```

### family_load_history.metadata
```json
{
    "duration_seconds": 1.5,
    "warnings": ["warning1"],
    "conflicts_resolved": 2
}
```
