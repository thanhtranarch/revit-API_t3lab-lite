"""
Vercel Serverless Function - Family Data API
Provides Revit family metadata for cloud-based family loading
"""

from http.server import BaseHTTPRequestHandler
import json


class handler(BaseHTTPRequestHandler):
    """Vercel serverless function handler for family data"""

    def do_GET(self):
        """Handle GET requests for family data"""

        # Sample family data structure
        # In production, this would come from a database or cloud storage
        families_data = {
            "categories": [
                {
                    "name": "Doors",
                    "path": "Architecture/Doors",
                    "families": [
                        {
                            "name": "Single-Flush Door",
                            "fileName": "Single-Flush.rfa",
                            "category": "Doors",
                            "size": 245760,
                            "downloadUrl": "https://your-storage.com/families/doors/Single-Flush.rfa",
                            "thumbnailUrl": "https://your-storage.com/thumbnails/doors/Single-Flush.png",
                            "description": "Standard single flush door",
                            "version": "2024",
                            "tags": ["door", "flush", "single"]
                        },
                        {
                            "name": "Double-Glass Door",
                            "fileName": "Double-Glass.rfa",
                            "category": "Doors",
                            "size": 312400,
                            "downloadUrl": "https://your-storage.com/families/doors/Double-Glass.rfa",
                            "thumbnailUrl": "https://your-storage.com/thumbnails/doors/Double-Glass.png",
                            "description": "Double door with glass panels",
                            "version": "2024",
                            "tags": ["door", "glass", "double"]
                        }
                    ]
                },
                {
                    "name": "Windows",
                    "path": "Architecture/Windows",
                    "families": [
                        {
                            "name": "Fixed Window",
                            "fileName": "Fixed.rfa",
                            "category": "Windows",
                            "size": 198500,
                            "downloadUrl": "https://your-storage.com/families/windows/Fixed.rfa",
                            "thumbnailUrl": "https://your-storage.com/thumbnails/windows/Fixed.png",
                            "description": "Standard fixed window",
                            "version": "2024",
                            "tags": ["window", "fixed"]
                        },
                        {
                            "name": "Casement Window",
                            "fileName": "Casement.rfa",
                            "category": "Windows",
                            "size": 225600,
                            "downloadUrl": "https://your-storage.com/families/windows/Casement.rfa",
                            "thumbnailUrl": "https://your-storage.com/thumbnails/windows/Casement.png",
                            "description": "Casement window with handle",
                            "version": "2024",
                            "tags": ["window", "casement", "operable"]
                        }
                    ]
                },
                {
                    "name": "Furniture",
                    "path": "Furniture/Office",
                    "families": [
                        {
                            "name": "Office Desk",
                            "fileName": "Office-Desk.rfa",
                            "category": "Furniture",
                            "size": 145800,
                            "downloadUrl": "https://your-storage.com/families/furniture/Office-Desk.rfa",
                            "thumbnailUrl": "https://your-storage.com/thumbnails/furniture/Office-Desk.png",
                            "description": "Standard office desk",
                            "version": "2024",
                            "tags": ["furniture", "desk", "office"]
                        },
                        {
                            "name": "Office Chair",
                            "fileName": "Office-Chair.rfa",
                            "category": "Furniture",
                            "size": 178900,
                            "downloadUrl": "https://your-storage.com/families/furniture/Office-Chair.rfa",
                            "thumbnailUrl": "https://your-storage.com/thumbnails/furniture/Office-Chair.png",
                            "description": "Ergonomic office chair",
                            "version": "2024",
                            "tags": ["furniture", "chair", "office", "seating"]
                        }
                    ]
                },
                {
                    "name": "Structural Columns",
                    "path": "Structure/Columns",
                    "families": [
                        {
                            "name": "Rectangular Column",
                            "fileName": "Rectangular-Column.rfa",
                            "category": "Structural Columns",
                            "size": 156700,
                            "downloadUrl": "https://your-storage.com/families/structure/Rectangular-Column.rfa",
                            "thumbnailUrl": "https://your-storage.com/thumbnails/structure/Rectangular-Column.png",
                            "description": "Rectangular structural column",
                            "version": "2024",
                            "tags": ["structure", "column", "rectangular"]
                        },
                        {
                            "name": "Round Column",
                            "fileName": "Round-Column.rfa",
                            "category": "Structural Columns",
                            "size": 142300,
                            "downloadUrl": "https://your-storage.com/families/structure/Round-Column.rfa",
                            "thumbnailUrl": "https://your-storage.com/thumbnails/structure/Round-Column.png",
                            "description": "Circular structural column",
                            "version": "2024",
                            "tags": ["structure", "column", "round", "circular"]
                        }
                    ]
                }
            ],
            "totalFamilies": 10,
            "lastUpdated": "2024-12-24T00:00:00Z"
        }

        # Set response headers
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')  # Enable CORS
        self.end_headers()

        # Send response
        self.wfile.write(json.dumps(families_data, indent=2).encode())

        return
