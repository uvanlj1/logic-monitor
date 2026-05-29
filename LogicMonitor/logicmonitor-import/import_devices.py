#!/usr/bin/env python3
"""
LogicMonitor Device Importer
Imports devices from CSV or Excel with support for custom properties and nested hostgroups.
Designed to run locally or as a manual GitHub Actions workflow.
"""

import os
import sys
import json
import time
import hmac
import base64
import hashlib
import logging
import argparse
import requests
import pandas as pd
import re
from typing import Dict, List, Any, Optional, Tuple
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    # python-dotenv is optional. GitHub Actions injects environment variables directly.
    pass

# Configuration class to manage settings
class Config:
    def __init__(self):
        # API Credentials - loaded from environment variables
        self.lm_access_id = os.environ.get("LM_ACCESS_ID")
        self.lm_access_key = os.environ.get("LM_ACCESS_KEY")
        self.lm_company = os.environ.get("LM_COMPANY")

        # Default settings
        self.logging_level = logging.INFO
        self.log_to_file = True
        self.log_to_stdout = True
        self.input_file_path = os.environ.get("LM_INPUT_PATH") or os.environ.get("LM_EXCEL_PATH", "devices.xlsx")
        self.api_call_delay = 1  # seconds between API calls
        self.dry_run = False  # Simulation mode
        self.update_existing = False  # Accepted for GitHub workflow compatibility; existing update logic can be added later
        self.quiet_mode = False  # Summary-only console output

        # Required columns (first 6 columns)
        self.ip_column = "IP"
        self.display_name_column = "DisplayName"
        self.host_group_column = "HostGroup"
        self.collector_group_id_column = "CollectorGroupID"
        self.auto_balanced_column = "Auto-Balanced"
        self.collector_id_column = "CollectorID"

        # Override with command line arguments
        self.parse_args()

        # Setup logging
        self.setup_logging()

        # Validate configuration
        self.validate()

    @staticmethod
    def str_to_bool(value):
        """Convert common CLI/GitHub Actions boolean strings to bool."""
        if isinstance(value, bool):
            return value
        if value is None:
            return False
        return str(value).strip().lower() in ["1", "true", "yes", "y", "on"]

    def parse_args(self):
        """Parse command line arguments and update configuration."""
        parser = argparse.ArgumentParser(
            description="Import devices into LogicMonitor from a CSV or Excel file"
        )

        # File inputs. --file and --csv are preferred for GitHub Actions.
        # --excel is retained so older local commands still work.
        parser.add_argument("--file", dest="input_file_path", help="Path to the CSV or Excel file")
        parser.add_argument("--csv", dest="input_file_path", help="Path to the CSV file")
        parser.add_argument("--excel", dest="input_file_path", help="Path to the Excel file")

        parser.add_argument("--company", dest="lm_company",
                           help="LogicMonitor company name, e.g. kineticit")
        parser.add_argument("--access-id", dest="lm_access_id",
                           help="LogicMonitor API access ID")
        parser.add_argument("--access-key", dest="lm_access_key",
                           help="LogicMonitor API access key")
        parser.add_argument("--delay", type=float, dest="api_call_delay",
                            help="Delay between API calls in seconds")
        parser.add_argument("--log-level", choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"],
                            help="Set logging level")
        parser.add_argument("--no-log-file", action="store_false", dest="log_to_file",
                            help="Disable logging to file")
        parser.add_argument("--no-log-stdout", action="store_false", dest="log_to_stdout",
                            help="Disable logging to stdout")

        # These accept both:
        #   --dry-run
        # and:
        #   --dry-run true
        # which is important because GitHub workflow boolean inputs arrive as strings.
        parser.add_argument("--dry-run", nargs="?", const="true", default=None,
                            help="true/false - simulate operations without making changes")
        parser.add_argument("--update-existing", nargs="?", const="true", default=None,
                            help="true/false - accepted for workflow compatibility")
        parser.add_argument("--quiet", action="store_true", dest="quiet_mode",
                            help="Display only summary information on console")

        args = parser.parse_args()

        # Update configuration with command line arguments
        for key, value in vars(args).items():
            if value is None:
                continue

            if key in ["dry_run", "update_existing"]:
                setattr(self, key, self.str_to_bool(value))
            elif key == "log_level":
                self.logging_level = getattr(logging, value)
            else:
                setattr(self, key, value)

    def setup_logging(self):
        """Configure logging based on settings."""
        log_format = '%(asctime)s [%(levelname)s] %(message)s'
        date_format = '%Y-%m-%d %H:%M:%S'

        # Reset handlers
        root_logger = logging.getLogger()
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)

        # Set root logger level
        root_logger.setLevel(self.logging_level)

        # Add file handler if enabled
        if self.log_to_file:
            log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "device_import.log")
            file_handler = logging.FileHandler(log_file)
            file_handler.setFormatter(logging.Formatter(log_format, date_format))
            root_logger.addHandler(file_handler)

        # Add stdout handler if enabled
        if self.log_to_stdout:
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setFormatter(logging.Formatter(log_format, date_format))

            # If in quiet mode, only show WARNING and above on console
            if self.quiet_mode:
                console_handler.setLevel(logging.WARNING)
            else:
                console_handler.setLevel(self.logging_level)

            root_logger.addHandler(console_handler)

        # Setup specific loggers
        logging.getLogger('requests').setLevel(logging.WARNING)
        logging.getLogger('urllib3').setLevel(logging.WARNING)

    def validate(self):
        """Validate configuration and print warnings."""
        logger = logging.getLogger(__name__)

        if not self.lm_access_id:
            logger.error("LM_ACCESS_ID not set. Use environment variable or --access-id")
            sys.exit(1)

        if not self.lm_access_key:
            logger.error("LM_ACCESS_KEY not set. Use environment variable or --access-key")
            sys.exit(1)

        if not self.lm_company:
            logger.error("LM_COMPANY not set. Use environment variable or --company")
            sys.exit(1)

        if not os.path.exists(self.input_file_path):
            logger.error(f"Input file not found: {self.input_file_path}")
            sys.exit(1)

        allowed_extensions = [".csv", ".xlsx", ".xls"]
        file_extension = os.path.splitext(self.input_file_path)[1].lower()
        if file_extension not in allowed_extensions:
            logger.error(
                f"Unsupported input file type: {file_extension}. "
                f"Supported types are: {', '.join(allowed_extensions)}"
            )
            sys.exit(1)

class DeviceImporter:
    """Imports devices from an Excel file into LogicMonitor."""

    # Class-level group cache (shared across all instances)
    _group_cache = {}

    def __init__(self, config):
        self.config = config
        self.logger = logging.getLogger(__name__)
        self.devices = []
        self.custom_columns = []
        # Initialize the summary logger
        self.setup_summary_logger()

    def clear_group_cache(self):
        """Clear the group cache."""
        DeviceImporter._group_cache.clear()
        self.logger.info("Group cache cleared")

    def setup_summary_logger(self):
        """Set up a special logger for summary information."""
        self.summary_logger = logging.getLogger('summary')
        self.summary_logger.setLevel(logging.INFO)

        # Remove existing handlers if any
        for handler in self.summary_logger.handlers[:]:
            self.summary_logger.removeHandler(handler)

        # Create console handler that always prints regardless of quiet mode
        console_handler = logging.StreamHandler(sys.stdout)
        formatter = logging.Formatter('%(asctime)s [SUMMARY] %(message)s', '%Y-%m-%d %H:%M:%S')
        console_handler.setFormatter(formatter)
        self.summary_logger.addHandler(console_handler)

        # Don't propagate to root logger to avoid duplicate messages
        self.summary_logger.propagate = False

    def direct_api_call(self, method: str, path: str, data=None) -> Tuple[int, Dict]:
        """Make an API call with optimized auth header generation."""
        url = f"https://{self.config.lm_company}.logicmonitor.com/santaba/rest{path}"

        # Get current time in milliseconds
        epoch = str(int(time.time() * 1000))

        # Data string if provided
        data_str = ""
        if data:
            # Debug logging for outgoing requests
            self.logger.debug(f"Original request data (Python dict): {data}")

            # IMPORTANT: Make sure we're not sending any default/null values in the JSON
            if method.upper() in ['POST', 'PUT']:
                # Remove any keys with None or 0 values
                if isinstance(data, dict):
                    data = {k: v for k, v in data.items() if v is not None and (not isinstance(v, int) or v != 0)}
                    self.logger.debug(f"Cleaned request data (Python dict): {data}")

            data_str = json.dumps(data)
            self.logger.debug(f"Request data as JSON string: {data_str}")

        # Extract resource path without query parameters
        resource_path = path.split('?')[0] if '?' in path else path

        # Concatenate Request details in the exact working pattern
        request_vars = method + epoch + data_str + resource_path

        # Construct signature
        digest = hmac.new(
            self.config.lm_access_key.encode(),
            msg=request_vars.encode(),
            digestmod=hashlib.sha256
        ).hexdigest()

        signature = base64.b64encode(digest.encode())

        # Construct headers with x-version: 3 to match Postman
        auth = f"LMv1 {self.config.lm_access_id}:{signature.decode()}:{epoch}"
        headers = {
            'Content-Type': 'application/json',
            'Authorization': auth,
            'x-version': '3'  # Added this header to match Postman request
        }

        self.logger.info(f"Making {method} request to {url}")

        # For dry run mode, log but don't execute actual POST/PUT/DELETE requests
        if self.config.dry_run and method.upper() in ['POST', 'PUT', 'PATCH', 'DELETE']:
            self.logger.info(f"DRY RUN: Would make {method} request to {url}")
            if data:
                self.logger.info(f"DRY RUN: Request data: {data_str}")
            # Return fake success for dry run
            return 200, {"status": 0, "data": {"id": 999}}

        # Apply API call delay
        time.sleep(self.config.api_call_delay)

        try:
            if method.upper() == 'GET':
                response = requests.get(url, headers=headers, timeout=30)
            elif method.upper() == 'POST':
                # Log the actual JSON being sent
                self.logger.info(f"Sending JSON payload: {data_str}")
                response = requests.post(url, data=data_str, headers=headers, timeout=30)
            elif method.upper() == 'PUT':
                response = requests.put(url, data=data_str, headers=headers, timeout=30)
            elif method.upper() == 'PATCH':
                response = requests.patch(url, data=data_str, headers=headers, timeout=30)
            elif method.upper() == 'DELETE':
                response = requests.delete(url, headers=headers, timeout=30)
            else:
                raise ValueError(f"Unsupported HTTP verb: {method}")

            self.logger.info(f"Response status code: {response.status_code}")

            if response.content:
                try:
                    response_data = response.json()
                    return response.status_code, response_data
                except json.JSONDecodeError:
                    self.logger.warning(f"Response is not valid JSON: {response.content}")
                    return response.status_code, {"content": response.content.decode('utf-8')}

            return response.status_code, {}
        except requests.exceptions.Timeout:
            self.logger.error(f"Request timed out: {url}")
            return 0, {"error": "Request timed out"}
        except requests.exceptions.RequestException as e:
            self.logger.error(f"API request failed: {e}")
            return 0, {"error": str(e)}
        except Exception as e:
            self.logger.error(f"Unexpected error in API call: {e}")
            return 0, {"error": str(e)}

    def get_group_by_name_and_parent(self, name: str, parent_id: int) -> Tuple[int, bool]:
        """Get a group by name and parent ID - optimized to avoid unnecessary lookups."""
        # First check the cache
        cache_key = f"{name}:{parent_id}"
        if cache_key in DeviceImporter._group_cache:
            cached_result = DeviceImporter._group_cache[cache_key]
            self.logger.info(f"Using cached result for group '{name}' (parent ID: {parent_id})")
            return cached_result

        # Use the more reliable filter method directly instead of trying both approaches
        path = f"/device/groups?filter=name:{name},parentId:{parent_id}"
        self.logger.info(f"Looking up group: name:{name}, parentId:{parent_id}")

        # Make API call without x-version header (which seems to work more reliably)
        url = f"https://{self.config.lm_company}.logicmonitor.com/santaba/rest{path}"
        epoch = str(int(time.time() * 1000))
        request_vars = "GET" + epoch + "" + path.split('?')[0]
        digest = hmac.new(
            self.config.lm_access_key.encode(),
            msg=request_vars.encode(),
            digestmod=hashlib.sha256
        ).hexdigest()
        signature = base64.b64encode(digest.encode())
        auth = f"LMv1 {self.config.lm_access_id}:{signature.decode()}:{epoch}"
        headers = {'Content-Type': 'application/json', 'Authorization': auth}

        self.logger.info(f"Making group lookup request to {url}")
        time.sleep(self.config.api_call_delay)

        try:
            response = requests.get(url, headers=headers, timeout=30)
            self.logger.info(f"Response status code: {response.status_code}")
            data = response.json() if response.content else {}

            if response.status_code == 200 and 'data' in data and 'items' in data['data']:
                if len(data['data']['items']) > 0:
                    group = data['data']['items'][0]
                    group_id = group['id']

                    applies_to = group.get('appliesTo', '')
                    self.logger.info(f"Group '{name}' appliesTo raw value: {repr(applies_to)}")
                    has_expression = bool(re.search(r'[a-zA-Z0-9_]+\s*[=!<>]+', applies_to))

                    if has_expression:
                        self.logger.info(f"Group '{name}' (ID: {group_id}) is a DYNAMIC group")
                        is_dynamic_group = True
                    else:
                        self.logger.info(f"Group '{name}' (ID: {group_id}) is a STATIC group")
                        is_dynamic_group = False

                    # Cache the result
                    result = (group_id, is_dynamic_group)
                    DeviceImporter._group_cache[cache_key] = result
                    return result
        except Exception as e:
            self.logger.error(f"Error while looking up group: {str(e)}")

        # Group not found
        self.logger.info(f"Group '{name}' not found with parent ID {parent_id}")
        return (-1, False)  # Default to static for new groups

    def create_group(self, name: str, parent_id: int) -> int:
        """Create a group in LogicMonitor."""
        # First check if group already exists
        result = self.get_group_by_name_and_parent(name, parent_id)
        group_id = result[0] if isinstance(result, tuple) else result

        if group_id > 0:
            self.logger.info(f"Group '{name}' already exists with ID {group_id}")
            return group_id

        self.logger.info(f"Creating group '{name}' with parent ID {parent_id}")

        # Group data - no description
        group_data = {
            "name": name,
            "parentId": parent_id
        }

        # Make the API call without x-version: 3 for group creation
        def api_call_without_version(method, path, data):
            url = f"https://{self.config.lm_company}.logicmonitor.com/santaba/rest{path}"
            epoch = str(int(time.time() * 1000))
            data_str = json.dumps(data) if data else ""
            request_vars = method + epoch + data_str + path.split('?')[0]
            digest = hmac.new(
                self.config.lm_access_key.encode(),
                msg=request_vars.encode(),
                digestmod=hashlib.sha256
            ).hexdigest()
            signature = base64.b64encode(digest.encode())
            auth = f"LMv1 {self.config.lm_access_id}:{signature.decode()}:{epoch}"
            headers = {'Content-Type': 'application/json', 'Authorization': auth}
            self.logger.info(f"Making {method} request without x-version to {url}")
            time.sleep(self.config.api_call_delay)
            response = requests.post(url, data=data_str, headers=headers, timeout=30) if method == "POST" else requests.get(url, headers=headers, timeout=30)
            self.logger.info(f"Response status code: {response.status_code}")
            return response.status_code, response.json() if response.content else {}

        status, data = api_call_without_version("POST", "/device/groups", group_data)

        if status == 200:
            # First attempt to extract ID directly from the response
            try:
                if "data" in data and "id" in data["data"]:
                    group_id = data["data"]["id"]
                    self.logger.info(f"Successfully created group '{name}' with ID {group_id}")
                    return group_id
                elif "id" in data:  # Check if ID is at the top level (like device responses)
                    group_id = data["id"]
                    self.logger.info(f"Successfully created group '{name}' with ID {group_id}")
                    return group_id
            except Exception as e:
                self.logger.warning(f"Could not extract group ID from response, will retry: {str(e)}")

            # If we couldn't extract the ID but got a 200 response, the group was likely created
            # Wait a moment to allow API operations to complete
            self.logger.info(f"Group '{name}' created successfully, but no ID returned. Waiting and looking it up...")
            time.sleep(2)  # Give the API time to register the new group

            # Try again to look up the group
            retry_result = self.get_group_by_name_and_parent(name, parent_id)
            if isinstance(retry_result, tuple):
                new_id = retry_result[0]
            else:
                new_id = retry_result

            if new_id > 0:
                self.logger.info(f"Successfully found newly created group '{name}' with ID {new_id}")
                return new_id

            self.logger.error(f"Group '{name}' was created but couldn't get ID")
            return -1
        elif status == 400 and "already exists" in json.dumps(data):
            # Group already exists, try to get it again
            self.logger.info(f"Group '{name}' already exists, trying to get ID")
            # Wait to ensure group is registered
            time.sleep(1)
            retry_result = self.get_group_by_name_and_parent(name, parent_id)
            if isinstance(retry_result, tuple):
                existing_id = retry_result[0]
            else:
                existing_id = retry_result

            if existing_id > 0:
                self.logger.info(f"Found existing group '{name}' with ID {existing_id}")
                return existing_id
            else:
                self.logger.error(f"Group '{name}' exists but couldn't retrieve ID")
                return -1
        else:
            error_msg = ""
            try:
                if isinstance(data, dict):
                    error_msg = data.get('errorMessage', data.get('errmsg', json.dumps(data)))
            except:
                error_msg = "Could not extract error message"

            self.logger.error(f"Failed to create group '{name}': {error_msg}")
            return -1

    def validate_collector_group(self, collector_group_id: int) -> bool:
         """Validate if collector group ID exists."""
         # For now, let's add a workaround to bypass this validation
         # since we're having API issues with the collector group endpoint
         self.logger.warning(f"WORKAROUND: Bypassing validation for collector group ID {collector_group_id}")
         return True

    def validate_collector(self, collector_id: int) -> bool:
        """Validate if collector ID exists."""
        # Skip validation if collector_id is 0 (auto-balanced) or invalid
        if collector_id <= 0:
            self.logger.info(f"Skipping validation for collector ID {collector_id} (auto-balanced or invalid)")
            return True

        resource_path = f'/setting/collector/collectors/{collector_id}'

        status, data = self.direct_api_call("GET", resource_path)

        if status == 200:
            return True
        else:
            self.logger.error(f"Collector with ID {collector_id} not found")
            return False

    def check_if_device_exists(self, ip_address: str, display_name: str) -> bool:
        """Check if a device with the given display name already exists."""
        resource_path = f'/device/devices'
        query_params = f'?filter=displayName:"{display_name}"'

        status, data = self.direct_api_call("GET", resource_path + query_params)

        # Check for direct items array as shown in the actual API response
        if status == 200 and isinstance(data, dict) and 'items' in data and len(data['items']) > 0:
            self.logger.info(f"Device with display name '{display_name}' already exists")
            return True

        # Fallback check for nested data.items structure
        if status == 200 and 'data' in data and 'items' in data['data'] and len(data['data']['items']) > 0:
            self.logger.info(f"Device with display name '{display_name}' already exists (nested format)")
            return True

        return False

    def add_device(self, device_data: Dict) -> bool:
        """Add a new device to LogicMonitor."""
        ip = device_data.get('ip')
        display_name = device_data.get('display_name')

        # Process host group path to get group ID
        group_path = device_data.get('host_group')

        # Skip all the group resolution if this is a re-run after failure
        # and we already know the final group ID
        final_group_id = device_data.get('resolved_group_id')
        leaf_is_dynamic = False

        if final_group_id is None:
            # Split the path into parts
            path_parts = [part.strip() for part in group_path.split('/') if part.strip()]
            if not path_parts:
                self.logger.error(f"Invalid group path format: {group_path}")
                return False

            # Start with root group
            parent_id = 1
            final_group_id = None

            # Navigate through the path parts to get to the final group
            for i, part in enumerate(path_parts):
                # Look up this group level with its parent
                result = self.get_group_by_name_and_parent(part, parent_id)

                # Extract group_id and is_dynamic from the result
                if isinstance(result, tuple) and len(result) == 2:
                    group_id, is_dynamic = result
                    self.logger.info(f"Group '{part}' check result: ID={group_id}, Dynamic={is_dynamic}")
                else:
                    # Fallback for unexpected result type
                    group_id = result if isinstance(result, int) else -1
                    is_dynamic = False
                    self.logger.info(f"Group '{part}' unexpected result type: {type(result)}")

                # If this is the leaf node, save its status
                if i == len(path_parts) - 1:
                    leaf_is_dynamic = is_dynamic
                    self.logger.info(f"Leaf node '{part}' dynamic status: {leaf_is_dynamic}")

                # If group doesn't exist, create it
                if group_id <= 0:
                    self.logger.info(f"Creating group '{part}' with parent ID {parent_id}")

                    # Create the group
                    group_data = {
                        "name": part,
                        "parentId": parent_id
                    }

                    # Make the request
                    status, create_data = self.direct_api_call("POST", "/device/groups", group_data)

                    if status == 200:
                        # First attempt to extract ID directly from the response
                        try:
                            if "data" in create_data and "id" in create_data["data"]:
                                group_id = create_data["data"]["id"]
                                self.logger.info(f"Successfully created group '{part}' with ID {group_id}")

                                # If this is the leaf node, it's a new group so not dynamic
                                if i == len(path_parts) - 1:
                                    leaf_is_dynamic = False
                            elif "id" in create_data:  # Check if ID is at the top level
                                group_id = create_data["id"]
                                self.logger.info(f"Successfully created group '{part}' with ID {group_id}")

                                # If this is the leaf node, it's a new group so not dynamic
                                if i == len(path_parts) - 1:
                                    leaf_is_dynamic = False
                        except Exception as e:
                            self.logger.warning(f"Could not extract group ID from response, will retry: {str(e)}")

                        # If we couldn't extract the ID but got a 200 response, the group was likely created
                        if group_id <= 0:
                            self.logger.info(f"Group '{part}' created successfully, but no ID returned. Waiting and looking it up...")
                            time.sleep(2)  # Give the API time to register the new group

                            # Try again to look up the group
                            retry_result = self.get_group_by_name_and_parent(part, parent_id)
                            if isinstance(retry_result, tuple):
                                group_id = retry_result[0]
                                is_dynamic = retry_result[1]
                            else:
                                group_id = retry_result

                            if group_id > 0:
                                self.logger.info(f"Successfully found newly created group '{part}' with ID {group_id}")
                                # If this is the leaf node, update its dynamic status
                                if i == len(path_parts) - 1:
                                    leaf_is_dynamic = is_dynamic
                    else:
                        error_msg = create_data.get('errorMessage', create_data.get('errmsg', "Unknown error"))
                        self.logger.error(f"Failed to create group '{part}' in path {group_path}: {error_msg}")

                        # If the error indicates group already exists, try to get it
                        if status == 400 and "already exists" in json.dumps(create_data):
                            self.logger.info(f"Group '{part}' already exists, trying to get ID")
                            # Wait to ensure group is registered
                            time.sleep(1)
                            retry_result = self.get_group_by_name_and_parent(part, parent_id)
                            if isinstance(retry_result, tuple):
                                group_id = retry_result[0]
                                is_dynamic = retry_result[1]
                            else:
                                group_id = retry_result

                            if group_id > 0:
                                self.logger.info(f"Successfully found existing group '{part}' with ID {group_id}")
                                # If this is the leaf node, update its dynamic status
                                if i == len(path_parts) - 1:
                                    leaf_is_dynamic = is_dynamic
                            else:
                                return False
                        else:
                            return False

                # This becomes the parent for the next iteration
                parent_id = group_id

                # If this is the last part (the leaf group), store its ID
                if i == len(path_parts) - 1:
                    final_group_id = group_id

            # Store the resolved group ID for potential future use
            device_data['resolved_group_id'] = final_group_id

        # Check if the final group was found
        if final_group_id is None:
            self.logger.error(f"Failed to resolve group path: {group_path}")
            self.summary_logger.error(f"Failed to add {display_name} - cannot resolve group path: {group_path}")
            return False

        # CRITICAL CHECK: If final group is dynamic, fail the operation
        self.logger.info(f"FINAL CHECK: Path={group_path}, ID={final_group_id}, Dynamic={leaf_is_dynamic}")
        if leaf_is_dynamic:
            self.logger.error(f"Cannot add device {display_name} to dynamic group {group_path}")
            self.logger.error(f"Dynamic groups are defined by 'appliesTo' expressions and not suitable for manual device addition")
            self.logger.error(f"Operation failed - skipping to next device")
            self.summary_logger.error(f"Failed to add {display_name} - target is a dynamic group")
            return False

        # We can proceed with a static group
        self.logger.info(f"Confirmed group {group_path} is STATIC - proceeding with device addition")

        # Get collector related info
        collector_group_id = device_data.get('collector_group_id')
        is_auto_balanced = device_data.get('is_auto_balanced')
        collector_id = device_data.get('collector_id')

        # Prepare custom properties
        custom_properties = []
        for key, value in device_data.get('custom_properties', {}).items():
            if value is not None and str(value).strip():
                custom_properties.append({
                    "name": key,
                    "value": str(value).strip()
                })

        # Debug log the custom properties
        self.logger.info(f"Adding {len(custom_properties)} custom properties to device {display_name}")

        # Create base device data
        api_data = {
            "name": ip,
            "displayName": display_name,
            "hostGroupIds": str(final_group_id),
            "customProperties": custom_properties
        }

        # Add collector settings based on auto-balanced flag
        if is_auto_balanced:
            # For auto-balanced, ONLY set autoBalancedCollectorGroupId
            self.logger.info(f"Auto-balanced is enabled, using autoBalancedCollectorGroupId: {collector_group_id}")
            api_data["autoBalancedCollectorGroupId"] = collector_group_id

            # ENSURE we're not accidentally setting collector_id to 0
            # Double-check the fields to be safe
            if "preferredCollectorId" in api_data:
                self.logger.warning("Removing unexpected preferredCollectorId from auto-balanced device")
                del api_data["preferredCollectorId"]
        else:
            # For specific collector, use preferredCollectorId and preferredCollectorGroupId
            self.logger.info(f"Using specific collector: {collector_id} in group: {collector_group_id}")

            # ONLY set these if collector_id is valid (greater than 0)
            if collector_id > 0:
                api_data["preferredCollectorId"] = collector_id
                api_data["preferredCollectorGroupId"] = collector_group_id

                # Only validate collector if not auto-balanced and if we haven't validated it before
                if not device_data.get('collector_validated', False):
                    if not self.validate_collector(collector_id):
                        return False
                    # Mark as validated to avoid repeated validation
                    device_data['collector_validated'] = True
            else:
                self.logger.error(f"Invalid collector ID {collector_id} for non-auto-balanced device")
                return False

        # Log the exact API data being sent
        self.logger.info(f"API endpoint: POST /device/devices")
        self.logger.info(f"API request data: {json.dumps(api_data, indent=2)}")

        self.logger.info(f"Adding device {display_name} ({ip}) to group {group_path} (ID: {final_group_id})")
        collector_msg = "Auto-balanced" if is_auto_balanced else api_data.get("preferredCollectorId", "Unknown")
        self.logger.info(f"Collector settings: GroupID={collector_group_id}, Auto-Balanced={is_auto_balanced}, CollectorID={collector_msg}")

        # MANUALLY VERIFY THE JSON BEFORE SENDING
        payload_str = json.dumps(api_data)

        # Make sure no "collector":0 or similar is hiding in the JSON using more precise pattern matching
        if '"collector": 0' in payload_str or '"collector":0' in payload_str:
            self.logger.error(f"ALERT: Found suspicious 'collector' with 0 value in payload")

        if '"collectorId": 0' in payload_str or '"collectorId":0' in payload_str:
            self.logger.error(f"ALERT: Found suspicious 'collectorId' with 0 value in payload")

        if '"preferredCollectorId": 0' in payload_str or '"preferredCollectorId":0' in payload_str:
            self.logger.error(f"ALERT: Found suspicious 'preferredCollectorId' with 0 value in payload")

        try:
            # Make the API call
            status, response = self.direct_api_call("POST", "/device/devices", api_data)

            # Log the full response for debugging
            self.logger.info(f"API response status: {status}")
            self.logger.info(f"API response: {json.dumps(response, indent=2)}")

            # Safely check if response and response["data"] exist
            if status == 200 and isinstance(response, dict):
                # Check for error status code in the response
                if "status" in response and response["status"] == 1404:
                    # This is the specific error we've been hitting
                    error_msg = response.get('errorMessage', response.get('errmsg', "Unknown error"))
                    self.logger.error(f"❌ Failed to add device {display_name} due to collector issue: {error_msg}")

                    # If we get the Collector(id=0) error, log it but don't automatically try fallback
                    if "Collector(id=0) does not exist" in json.dumps(response) and is_auto_balanced:
                        self.logger.error("The API is reporting a Collector(id=0) error even though we're requesting auto-balanced.")
                        self.logger.error("This might be a LogicMonitor API issue with auto-balanced devices.")
                        self.summary_logger.error(f"Failed to add {display_name} - collector issue: {error_msg}")

                    return False

                # FIX: Check for ID directly at top level of response (matches actual API response)
                if "id" in response:
                    device_id = response["id"]
                    self.logger.info(f"✅ Successfully added device {display_name} (ID: {device_id})")
                    self.summary_logger.info(f"Successfully added {display_name} to {group_path}")
                    return True
                # Original checks for response.data.id structure
                elif "data" in response and isinstance(response["data"], dict) and "id" in response["data"]:
                    device_id = response["data"]["id"]
                    self.logger.info(f"✅ Successfully added device {display_name} (ID: {device_id})")
                    self.summary_logger.info(f"Successfully added {display_name} to {group_path}")
                    return True
                elif response.get("status") == 0:  # Some APIs return status 0 for success without data
                    self.logger.info(f"✅ Successfully added device {display_name} (no ID returned)")
                    self.summary_logger.info(f"Successfully added {display_name} to {group_path}")
                    return True
                else:
                    # The fix for the issue in the logs: Despite error message, device might be created
                    # Let's check if the device exists after the API call
                    self.logger.warning(f"API returned success status but no ID was found in response")
                    self.logger.warning(f"Checking if device was actually created despite parsing issue...")

                    # Wait a moment to allow API operations to complete
                    time.sleep(2)

                    # Check if device exists now
                    if self.check_if_device_exists(ip, display_name):
                        self.logger.info(f"✅ Device {display_name} exists in LogicMonitor - operation was successful despite response parsing issue")
                        self.summary_logger.info(f"Successfully added {display_name} to {group_path} (with response issues)")
                        return True
                    else:
                        error_msg = response.get('errorMessage', response.get('errmsg', "Unknown error"))
                        self.logger.error(f"Device {display_name} was not created. API returned success status but invalid data: {error_msg}")
                        self.summary_logger.error(f"Failed to add {display_name} - API error: {error_msg}")
                        return False
            else:
                error_msg = ""
                if isinstance(response, dict):
                    error_msg = response.get('errorMessage', response.get('errmsg', "Unknown error"))
                self.logger.error(f"❌ Failed to add device {display_name}: {error_msg}")
                self.summary_logger.error(f"Failed to add {display_name}: {error_msg}")
                return False
        except Exception as e:
            self.logger.error(f"❌ Unexpected error adding device {display_name}: {str(e)}")
            self.summary_logger.error(f"Failed to add {display_name} - unexpected error: {str(e)}")
            # Check if the device was created despite the error
            time.sleep(2)
            if self.check_if_device_exists(ip, display_name):
                self.logger.info(f"✅ Device {display_name} exists in LogicMonitor despite error - operation was successful")
                self.summary_logger.info(f"Successfully added {display_name} despite errors")
                return True
            return False

    def load_input_file(self) -> bool:
        """Load device definitions from a CSV or Excel file."""
        try:
            self.logger.info(f"Reading input file: {self.config.input_file_path}")
            file_extension = os.path.splitext(self.config.input_file_path)[1].lower()

            if file_extension == ".csv":
                df = pd.read_csv(self.config.input_file_path)
            elif file_extension in [".xlsx", ".xls"]:
                df = pd.read_excel(self.config.input_file_path)
            else:
                self.logger.error(f"Unsupported input file type: {file_extension}")
                return False

            # Check required columns
            required_columns = [
                self.config.ip_column,
                self.config.display_name_column,
                self.config.host_group_column,
                self.config.collector_group_id_column,
                self.config.auto_balanced_column,
                self.config.collector_id_column
            ]

            for column in required_columns:
                if column not in df.columns:
                    self.logger.error(f"Required column '{column}' not found in input file")
                    return False

            # Identify custom property columns (all columns except required ones)
            self.custom_columns = [col for col in df.columns if col not in required_columns]
            self.logger.info(f"Found {len(self.custom_columns)} custom property columns: {', '.join(self.custom_columns)}")

            # Process each row
            for index, row in df.iterrows():
                # Basic validation
                ip = str(row[self.config.ip_column]).strip()
                display_name = str(row[self.config.display_name_column]).strip()
                host_group = str(row[self.config.host_group_column]).strip()

                try:
                    collector_group_id = int(row[self.config.collector_group_id_column])
                except (ValueError, TypeError):
                    self.logger.error(f"Row {index+2}: Invalid collector group ID: {row[self.config.collector_group_id_column]}")
                    continue

                # Parse auto-balanced value (case-insensitive)
                auto_balanced_str = str(row[self.config.auto_balanced_column]).strip().lower()
                is_auto_balanced = auto_balanced_str in ['yes', 'y', 'true', '1']
                self.logger.info(f"Row {index+2}: Auto-balanced value: '{auto_balanced_str}' interpreted as: {is_auto_balanced}")

                # Collector ID handling
                collector_id = 0  # Default for auto-balanced
                if not is_auto_balanced:
                    try:
                        collector_id = int(row[self.config.collector_id_column])
                        if collector_id <= 0:
                            self.logger.error(f"Row {index+2}: Invalid collector ID: {collector_id} for non-auto-balanced device")
                            continue
                    except (ValueError, TypeError):
                        self.logger.error(f"Row {index+2}: Invalid collector ID: {row[self.config.collector_id_column]}")
                        continue

                if not ip:
                    self.logger.warning(f"Row {index+2}: Missing IP, skipping")
                    continue

                if not display_name:
                    self.logger.warning(f"Row {index+2}: Missing display name, skipping")
                    continue

                if not host_group:
                    self.logger.warning(f"Row {index+2}: Missing host group, skipping")
                    continue

                # Collect custom properties
                custom_properties = {}
                for col in self.custom_columns:
                    if pd.notna(row[col]):  # Check if value is not NaN
                        custom_properties[col] = row[col]

                # Create device data dictionary
                device = {
                    'ip': ip,
                    'display_name': display_name,
                    'host_group': host_group,
                    'collector_group_id': collector_group_id,
                    'is_auto_balanced': is_auto_balanced,
                    'collector_id': collector_id,
                    'custom_properties': custom_properties
                }

                self.devices.append(device)

            self.logger.info(f"Loaded {len(self.devices)} devices from input file")
            return True

        except Exception as e:
            self.logger.error(f"Error loading input file: {e}")
            return False

    def import_devices(self) -> Tuple[int, int, int]:
        """Import all devices from the Excel file."""
        success_count = 0
        skipped_count = 0
        error_count = 0

        # Use both loggers, normal logger for details and summary for important info
        self.logger.info(f"Starting import of {len(self.devices)} devices")
        self.summary_logger.info(f"Starting import of {len(self.devices)} devices")

        if self.config.dry_run:
            self.logger.info("=== DRY RUN MODE - No devices will be added ===")
            self.summary_logger.info("=== DRY RUN MODE - No devices will be added ===")

        # Clear the group cache at the start of a batch import
        # This ensures we have a clean slate but don't clear between devices
        DeviceImporter._group_cache.clear()

        for i, device in enumerate(self.devices):
            ip = device['ip']
            display_name = device['display_name']

            self.logger.info(f"Processing device {i+1}/{len(self.devices)}: {display_name} ({ip})")
            self.summary_logger.info(f"Processing {i+1}/{len(self.devices)}: {display_name}")

            # First, explicitly check if device already exists
            if self.check_if_device_exists(ip, display_name):
                self.logger.info(f"Device with display name '{display_name}' already exists, skipping")
                self.summary_logger.info(f"Skipping {display_name} - already exists")
                skipped_count += 1
                continue  # Skip to next device

            # Process the device
            try:
                result = self.add_device(device)
                if result:
                    success_count += 1
                else:
                    # Double-check if device exists despite the add_device failure
                    # This can happen if the API behavior is inconsistent
                    if self.check_if_device_exists(ip, display_name):
                        self.logger.info(f"Device {display_name} appears to have been created despite error, marking as success")
                        self.summary_logger.info(f"Device {display_name} created despite errors")
                        success_count += 1
                    else:
                        error_count += 1
            except Exception as e:
                self.logger.error(f"Unexpected error processing device {display_name}: {str(e)}")
                self.summary_logger.error(f"Error adding {display_name}: {str(e)}")
                error_count += 1

        # Print summary
        summary_text = [
            "=" * 50,
            "Import Summary:",
            f"  Total devices processed: {len(self.devices)}",
            f"  Successfully added: {success_count}",
            f"  Skipped (already exist): {skipped_count}",
            f"  Failed: {error_count}"
        ]

        # Log to both loggers
        for line in summary_text:
            self.logger.info(line)
            self.summary_logger.info(line)

        if self.config.dry_run:
            self.logger.info("=== DRY RUN MODE - No actual changes were made ===")
            self.summary_logger.info("=== DRY RUN MODE - No actual changes were made ===")

        return success_count, skipped_count, error_count
def main():
    """Main entry point for the script."""
    try:
        # Initialize configuration
        config = Config()
        logger = logging.getLogger(__name__)

        if config.quiet_mode:
            summary_logger = logging.getLogger('summary')
            summary_logger.info("Running in QUIET mode - detailed logs saved to device_import.log")
        else:
            logger.info("Starting LogicMonitor Device Importer")
            logger.info(f"Using input file: {config.input_file_path}")

        if config.dry_run:
            logger.info("Running in DRY RUN mode - no devices will be added")

        # Initialize importer
        importer = DeviceImporter(config)

        # Load input data
        if not importer.load_input_file():
            logger.error("Failed to load input data. Exiting.")
            return 1

        # Import devices
        success_count, skipped_count, error_count = importer.import_devices()

        return 0 if error_count == 0 else 1

    except KeyboardInterrupt:
        logger.info("Operation interrupted by user")
        return 130
    except Exception as e:
        logger.exception(f"Unhandled exception: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
