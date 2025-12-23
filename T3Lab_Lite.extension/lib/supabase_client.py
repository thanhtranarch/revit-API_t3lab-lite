# -*- coding: utf-8 -*-
"""
Supabase REST API Client for IronPython
Provides integration with Supabase database for Revit family management
"""
__title__ = "Supabase Client"
__author__ = "T3Lab"

import os
import sys
import clr
import json
import hashlib
import traceback
from datetime import datetime

# .NET Imports for HTTP requests
clr.AddReference("System")
clr.AddReference("System.Net")
clr.AddReference("System.Web")
from System import Uri
from System.Net import WebClient, WebRequest, WebException
from System.Text import Encoding
from System.IO import StreamReader

# pyRevit Imports
try:
    from pyrevit import script
    logger = script.get_logger()
except:
    # Fallback if pyRevit is not available
    import logging
    logger = logging.getLogger(__name__)

# Configuration
CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".t3lab")
SUPABASE_CONFIG_FILE = os.path.join(CONFIG_DIR, "supabase_config.json")


class SupabaseClient:
    """REST API client for Supabase using .NET WebClient"""

    def __init__(self, url=None, api_key=None, bucket_name='revit-family-free'):
        """
        Initialize Supabase client

        Args:
            url: Supabase project URL (e.g., https://xxxxx.supabase.co)
            api_key: Supabase anon/public API key
            bucket_name: Storage bucket name (default: revit-family-free)
        """
        self.url = url
        self.api_key = api_key
        self.bucket_name = bucket_name
        self.rest_url = None
        self.storage_url = None

        if url:
            self.rest_url = "{}/rest/v1".format(url.rstrip('/'))
            self.storage_url = "{}/storage/v1".format(url.rstrip('/'))

        # Load config if no credentials provided
        if not self.url or not self.api_key:
            self.load_config()

    def load_config(self):
        """Load Supabase configuration from file"""
        try:
            if os.path.exists(SUPABASE_CONFIG_FILE):
                with open(SUPABASE_CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                    self.url = config.get('url')
                    self.api_key = config.get('api_key')
                    self.bucket_name = config.get('bucket_name', 'revit-family-free')
                    if self.url:
                        self.rest_url = "{}/rest/v1".format(self.url.rstrip('/'))
                        self.storage_url = "{}/storage/v1".format(self.url.rstrip('/'))
                    logger.info("Supabase config loaded from: {}".format(SUPABASE_CONFIG_FILE))
                    return True
            else:
                logger.warning("No Supabase config file found at: {}".format(SUPABASE_CONFIG_FILE))
                return False
        except Exception as ex:
            logger.error("Failed to load Supabase config: {}".format(ex))
            logger.error(traceback.format_exc())
            return False

    def save_config(self, url, api_key, bucket_name='revit-family-free'):
        """Save Supabase configuration to file"""
        try:
            # Create config directory if needed
            if not os.path.exists(CONFIG_DIR):
                os.makedirs(CONFIG_DIR)

            config = {
                'url': url,
                'api_key': api_key,
                'bucket_name': bucket_name
            }

            with open(SUPABASE_CONFIG_FILE, 'w') as f:
                json.dump(config, f, indent=4)

            self.url = url
            self.api_key = api_key
            self.bucket_name = bucket_name
            self.rest_url = "{}/rest/v1".format(url.rstrip('/'))
            self.storage_url = "{}/storage/v1".format(url.rstrip('/'))

            logger.info("Supabase config saved to: {}".format(SUPABASE_CONFIG_FILE))
            return True
        except Exception as ex:
            logger.error("Failed to save Supabase config: {}".format(ex))
            logger.error(traceback.format_exc())
            return False

    def is_configured(self):
        """Check if client is properly configured"""
        return bool(self.url and self.api_key and self.rest_url)

    def _make_request(self, method, endpoint, data=None, params=None):
        """
        Make HTTP request to Supabase REST API

        Args:
            method: HTTP method (GET, POST, PATCH, DELETE)
            endpoint: API endpoint (e.g., 'families')
            data: Request body data (dict)
            params: Query parameters (dict)

        Returns:
            Response data (dict or list)
        """
        if not self.is_configured():
            raise Exception("Supabase client not configured. Please set URL and API key.")

        try:
            # Build URL with query parameters
            url = "{}/{}".format(self.rest_url, endpoint)
            if params:
                query_parts = []
                for key, value in params.items():
                    query_parts.append("{}={}".format(key, value))
                url = "{}?{}".format(url, "&".join(query_parts))

            logger.debug("Supabase request: {} {}".format(method, url))

            # Create web client
            client = WebClient()

            # Set headers
            client.Headers.Add("apikey", self.api_key)
            client.Headers.Add("Authorization", "Bearer {}".format(self.api_key))
            client.Headers.Add("Content-Type", "application/json")
            client.Headers.Add("Prefer", "return=representation")

            # Make request
            response_text = None

            if method == "GET":
                response_text = client.DownloadString(url)

            elif method == "POST":
                json_data = json.dumps(data) if data else "{}"
                response_text = client.UploadString(url, "POST", json_data)

            elif method == "PATCH":
                json_data = json.dumps(data) if data else "{}"
                response_text = client.UploadString(url, "PATCH", json_data)

            elif method == "DELETE":
                client.UploadString(url, "DELETE", "")
                return None

            else:
                raise Exception("Unsupported HTTP method: {}".format(method))

            # Parse response
            if response_text:
                result = json.loads(response_text)
                logger.debug("Supabase response: {} records".format(
                    len(result) if isinstance(result, list) else 1
                ))
                return result

            return None

        except WebException as web_ex:
            # Try to read error response
            error_msg = str(web_ex)
            try:
                if web_ex.Response:
                    stream = web_ex.Response.GetResponseStream()
                    reader = StreamReader(stream)
                    error_text = reader.ReadToEnd()
                    error_msg = "HTTP Error: {}".format(error_text)
            except:
                pass

            logger.error("Supabase request failed: {}".format(error_msg))
            raise Exception("Supabase API error: {}".format(error_msg))

        except Exception as ex:
            logger.error("Supabase request exception: {}".format(ex))
            logger.error(traceback.format_exc())
            raise

    # Family Management Methods

    def insert_family(self, name, file_path, category=None, file_size=None,
                     project_name=None, user_name=None, metadata=None):
        """
        Insert a new family record

        Args:
            name: Family name
            file_path: Full path to family file
            category: Family category
            file_size: File size in bytes
            project_name: Current Revit project name
            user_name: Current user name
            metadata: Additional metadata (dict)

        Returns:
            Created family record (dict)
        """
        try:
            # Calculate file hash for deduplication
            file_hash = None
            if os.path.exists(file_path):
                file_hash = self._calculate_file_hash(file_path)
                if not file_size:
                    file_size = os.path.getsize(file_path)

            data = {
                'name': name,
                'file_path': file_path,
                'category': category,
                'file_size': file_size,
                'file_hash': file_hash,
                'project_name': project_name,
                'user_name': user_name,
                'metadata': metadata or {}
            }

            result = self._make_request('POST', 'families', data=data)
            logger.info("Inserted family: {}".format(name))
            return result[0] if isinstance(result, list) else result

        except Exception as ex:
            logger.error("Failed to insert family {}: {}".format(name, ex))
            raise

    def get_family_by_name(self, name):
        """
        Get family by name

        Args:
            name: Family name

        Returns:
            Family record (dict) or None
        """
        try:
            params = {'name': 'eq.{}'.format(name)}
            result = self._make_request('GET', 'families', params=params)

            if result and len(result) > 0:
                return result[0]
            return None

        except Exception as ex:
            logger.error("Failed to get family {}: {}".format(name, ex))
            return None

    def get_family_by_hash(self, file_hash):
        """
        Get family by file hash

        Args:
            file_hash: SHA256 hash of file

        Returns:
            Family record (dict) or None
        """
        try:
            params = {'file_hash': 'eq.{}'.format(file_hash)}
            result = self._make_request('GET', 'families', params=params)

            if result and len(result) > 0:
                return result[0]
            return None

        except Exception as ex:
            logger.error("Failed to get family by hash: {}".format(ex))
            return None

    def update_family(self, family_id, **updates):
        """
        Update family record

        Args:
            family_id: Family UUID
            **updates: Fields to update

        Returns:
            Updated family record (dict)
        """
        try:
            # Add updated_at timestamp
            updates['updated_at'] = datetime.utcnow().isoformat()

            params = {'id': 'eq.{}'.format(family_id)}
            result = self._make_request('PATCH', 'families', data=updates, params=params)

            logger.info("Updated family: {}".format(family_id))
            return result[0] if isinstance(result, list) else result

        except Exception as ex:
            logger.error("Failed to update family {}: {}".format(family_id, ex))
            raise

    def get_all_families(self, order_by='name', limit=None):
        """
        Get all families

        Args:
            order_by: Field to order by (default: name)
            limit: Maximum records to return

        Returns:
            List of family records
        """
        try:
            params = {
                'select': '*',
                'order': '{}.asc'.format(order_by)
            }

            if limit:
                params['limit'] = str(limit)

            result = self._make_request('GET', 'families', params=params)
            return result or []

        except Exception as ex:
            logger.error("Failed to get all families: {}".format(ex))
            return []

    def search_families(self, search_term, category=None):
        """
        Search families by name or category

        Args:
            search_term: Search term
            category: Filter by category

        Returns:
            List of matching family records
        """
        try:
            params = {
                'select': '*',
                'or': '(name.ilike.*{}*,category.ilike.*{}*)'.format(search_term, search_term),
                'order': 'name.asc'
            }

            if category:
                params['category'] = 'eq.{}'.format(category)

            result = self._make_request('GET', 'families', params=params)
            return result or []

        except Exception as ex:
            logger.error("Failed to search families: {}".format(ex))
            return []

    # Load History Methods

    def insert_load_history(self, family_id, project_name, user_name,
                           success=True, error_message=None,
                           revit_version=None, metadata=None):
        """
        Insert load history record

        Args:
            family_id: Family UUID
            project_name: Revit project name
            user_name: User name
            success: Whether load was successful
            error_message: Error message if failed
            revit_version: Revit version
            metadata: Additional metadata

        Returns:
            Created load history record (dict)
        """
        try:
            data = {
                'family_id': family_id,
                'project_name': project_name,
                'user_name': user_name,
                'success': success,
                'error_message': error_message,
                'revit_version': revit_version,
                'metadata': metadata or {}
            }

            result = self._make_request('POST', 'family_load_history', data=data)
            logger.info("Inserted load history for family: {}".format(family_id))
            return result[0] if isinstance(result, list) else result

        except Exception as ex:
            logger.error("Failed to insert load history: {}".format(ex))
            raise

    def get_load_history(self, family_id=None, project_name=None, limit=50):
        """
        Get load history

        Args:
            family_id: Filter by family UUID
            project_name: Filter by project name
            limit: Maximum records to return

        Returns:
            List of load history records
        """
        try:
            params = {
                'select': '*',
                'order': 'loaded_at.desc',
                'limit': str(limit)
            }

            if family_id:
                params['family_id'] = 'eq.{}'.format(family_id)

            if project_name:
                params['project_name'] = 'eq.{}'.format(project_name)

            result = self._make_request('GET', 'family_load_history', params=params)
            return result or []

        except Exception as ex:
            logger.error("Failed to get load history: {}".format(ex))
            return []

    # Storage Methods

    def upload_family_file(self, file_path, storage_path=None):
        """
        Upload family file to Supabase storage

        Args:
            file_path: Local file path
            storage_path: Path in storage bucket (default: filename only)

        Returns:
            Storage path in bucket
        """
        if not self.is_configured():
            raise Exception("Supabase client not configured")

        try:
            # Default storage path is just the filename
            if not storage_path:
                storage_path = os.path.basename(file_path)

            # Build upload URL
            upload_url = "{}/object/{}/{}".format(
                self.storage_url,
                self.bucket_name,
                storage_path
            )

            logger.info("Uploading file to: {}".format(upload_url))

            # Create web client
            client = WebClient()
            client.Headers.Add("apikey", self.api_key)
            client.Headers.Add("Authorization", "Bearer {}".format(self.api_key))
            client.Headers.Add("Content-Type", "application/octet-stream")

            # Upload file
            response = client.UploadFile(upload_url, "POST", file_path)
            response_text = Encoding.UTF8.GetString(response)

            logger.info("File uploaded successfully: {}".format(storage_path))
            return storage_path

        except WebException as web_ex:
            error_msg = str(web_ex)
            try:
                if web_ex.Response:
                    stream = web_ex.Response.GetResponseStream()
                    reader = StreamReader(stream)
                    error_text = reader.ReadToEnd()
                    error_msg = "Upload Error: {}".format(error_text)
            except:
                pass

            logger.error("Failed to upload file: {}".format(error_msg))
            raise Exception("Supabase storage upload error: {}".format(error_msg))

        except Exception as ex:
            logger.error("Upload exception: {}".format(ex))
            logger.error(traceback.format_exc())
            raise

    def download_family_file(self, storage_path, local_path):
        """
        Download family file from Supabase storage

        Args:
            storage_path: Path in storage bucket
            local_path: Local destination path

        Returns:
            Local file path
        """
        if not self.is_configured():
            raise Exception("Supabase client not configured")

        try:
            # Build download URL
            download_url = "{}/object/{}/{}".format(
                self.storage_url,
                self.bucket_name,
                storage_path
            )

            logger.info("Downloading file from: {}".format(download_url))

            # Create web client
            client = WebClient()
            client.Headers.Add("apikey", self.api_key)
            client.Headers.Add("Authorization", "Bearer {}".format(self.api_key))

            # Download file
            client.DownloadFile(download_url, local_path)

            logger.info("File downloaded successfully: {}".format(local_path))
            return local_path

        except WebException as web_ex:
            error_msg = str(web_ex)
            try:
                if web_ex.Response:
                    stream = web_ex.Response.GetResponseStream()
                    reader = StreamReader(stream)
                    error_text = reader.ReadToEnd()
                    error_msg = "Download Error: {}".format(error_text)
            except:
                pass

            logger.error("Failed to download file: {}".format(error_msg))
            raise Exception("Supabase storage download error: {}".format(error_msg))

        except Exception as ex:
            logger.error("Download exception: {}".format(ex))
            logger.error(traceback.format_exc())
            raise

    def get_public_url(self, storage_path):
        """
        Get public URL for a file in storage

        Args:
            storage_path: Path in storage bucket

        Returns:
            Public URL
        """
        if not self.is_configured():
            return None

        return "{}/object/public/{}/{}".format(
            self.storage_url,
            self.bucket_name,
            storage_path
        )

    def list_storage_files(self, folder_path=''):
        """
        List files in storage bucket

        Args:
            folder_path: Folder path in bucket (optional)

        Returns:
            List of file objects
        """
        if not self.is_configured():
            raise Exception("Supabase client not configured")

        try:
            # Build list URL
            list_url = "{}/object/list/{}".format(
                self.storage_url,
                self.bucket_name
            )

            # Add folder path if provided
            data = {"prefix": folder_path} if folder_path else {}

            # Create web client
            client = WebClient()
            client.Headers.Add("apikey", self.api_key)
            client.Headers.Add("Authorization", "Bearer {}".format(self.api_key))
            client.Headers.Add("Content-Type", "application/json")

            # Make request
            json_data = json.dumps(data)
            response_text = client.UploadString(list_url, "POST", json_data)

            # Parse response
            result = json.loads(response_text)
            logger.info("Listed {} files in storage".format(len(result)))
            return result

        except Exception as ex:
            logger.error("Failed to list storage files: {}".format(ex))
            logger.error(traceback.format_exc())
            return []

    # Utility Methods

    def _calculate_file_hash(self, file_path):
        """Calculate SHA256 hash of file"""
        try:
            sha256 = hashlib.sha256()
            with open(file_path, 'rb') as f:
                while True:
                    chunk = f.read(8192)
                    if not chunk:
                        break
                    sha256.update(chunk)
            return sha256.hexdigest()
        except Exception as ex:
            logger.error("Failed to calculate file hash: {}".format(ex))
            return None

    def sync_family(self, name, file_path, category=None, project_name=None, user_name=None, upload_file=True):
        """
        Sync family to Supabase (insert or update) with optional file upload

        Args:
            name: Family name
            file_path: Full path to family file
            category: Family category
            project_name: Current project name
            user_name: Current user name
            upload_file: Whether to upload file to storage (default: True)

        Returns:
            Family record (dict) with 'storage_path' if uploaded
        """
        try:
            # Check if family exists by hash
            file_hash = self._calculate_file_hash(file_path)
            existing = self.get_family_by_hash(file_hash) if file_hash else None

            # Upload file to storage if requested
            storage_path = None
            if upload_file and os.path.exists(file_path):
                try:
                    # Create storage path: category/filename
                    filename = os.path.basename(file_path)
                    if category:
                        storage_path = "{}/{}".format(category.replace('\\', '/'), filename)
                    else:
                        storage_path = filename

                    logger.info("Uploading family file to storage: {}".format(storage_path))
                    self.upload_family_file(file_path, storage_path)
                    logger.info("File uploaded successfully")
                except Exception as upload_ex:
                    logger.warning("Failed to upload file to storage: {}".format(upload_ex))
                    # Continue with database sync even if upload fails
                    storage_path = None

            # Prepare metadata
            metadata = {
                'storage_path': storage_path
            } if storage_path else {}

            if existing:
                # Update existing family
                logger.info("Family exists, updating: {}".format(name))
                updates = {
                    'name': name,
                    'file_path': file_path,
                    'category': category,
                    'load_count': existing.get('load_count', 0) + 1,
                    'last_loaded_at': datetime.utcnow().isoformat()
                }
                if storage_path:
                    existing_metadata = existing.get('metadata', {})
                    existing_metadata['storage_path'] = storage_path
                    updates['metadata'] = existing_metadata

                result = self.update_family(existing['id'], **updates)
            else:
                # Insert new family
                logger.info("New family, inserting: {}".format(name))
                file_size = os.path.getsize(file_path) if os.path.exists(file_path) else None
                result = self.insert_family(
                    name=name,
                    file_path=file_path,
                    category=category,
                    file_size=file_size,
                    project_name=project_name,
                    user_name=user_name,
                    metadata=metadata
                )

            # Add storage_path to result
            if storage_path and result:
                result['storage_path'] = storage_path

            return result

        except Exception as ex:
            logger.error("Failed to sync family {}: {}".format(name, ex))
            raise


# Singleton instance
_client = None

def get_client():
    """Get or create singleton Supabase client instance"""
    global _client
    if _client is None:
        _client = SupabaseClient()
    return _client
