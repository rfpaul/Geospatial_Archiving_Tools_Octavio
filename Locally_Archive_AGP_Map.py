#!/usr/bin/env python3
"""
Locally_Archive_AGP_Map.py

Creates a timestamped directory with a local copy of all the layers in an
ArcGIS Pro Map as a file geodatabase. The script will also attempt to preserve
metadata and attach it to the data in the file geodatabase.

* Copyright 2025 The Board of Trustees of the University of Illinois. All Rights Reserved.
* Licensed under the terms of the GNU General Public License version 3.
* The License is included in the distribution as License.txt file.
* You may not use this file except in compliance with the License.
* Software distributed under the License is distributed on an "AS IS" BASIS,
* WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
* See the License for the specific language governing permissions and limitations under the Option.

Created by Robert Paul at the Illinois State Geological Survey (ISGS) with
assistance from Claude (Sonnet 3.5)
"""

import arcpy
import os
import shutil
import datetime
import subprocess
import re
from arcpy import metadata as md

def get_layer_extent(map_obj, layer_name):
    """
    Gets the extent of a specified layer in a map using Describe.
    
    Args:
        map_obj: ArcGIS Pro Map object
        layer_name: Name of the layer to get extent from
    
    Returns:
        arcpy.Extent object or None if extent cannot be determined
        
    Raises:
        ValueError: If the specified layer is not found or has no extent
    """
    for layer in map_obj.listLayers():
        if layer.name == layer_name and not layer.isGroupLayer:
            desc = arcpy.Describe(layer)
            if hasattr(desc, 'extent'):
                return desc.extent
            else:
                raise ValueError(f"Layer '{layer_name}' has no spatial extent")
    
    raise ValueError(f"Layer '{layer_name}' not found in map or is a group layer")

def find_7zip_path():
    """
    Finds the path to 7-Zip executable on Windows.
    
    Returns:
        Path to 7z.exe
        
    Raises:
        ValueError: If 7-Zip is not found
    """
    potential_paths = [
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe"
    ]
    
    for path in potential_paths:
        if os.path.exists(path):
            return path
            
    raise ValueError("7-Zip not found. Please install 7-Zip from https://7-zip.org")

def package_map(aprx, map_obj, output_path, extent=None):
    """
    Packages a map to an MPKX file, handling Z-enabled layers separately.
    
    Args:
        aprx: ArcGIS Pro Project connection
        map_obj: ArcGIS Pro Map object
        output_path: Path for the output MPKX file
        extent: Optional extent to limit the data
        
    Returns:
        Path to the created MPKX file
    """
    print(f"Packaging map to {output_path}...")
    map_name = arcpy.ValidateTableName(map_obj.name)
    mpkx_path = os.path.join(output_path, f"{map_name}.mpkx")
    
    # Create temporary map for non-Z layers
    temp_map = aprx.createMap("Temp Map")
    
    try:
        # Identify Z and non-Z layers
        non_z_layers, z_layers = identify_z_layers(map_obj)
        
        if z_layers:
            print(f"Found {len(z_layers)} layers with Z values that will be processed separately")
        
        # Add non-Z layers to temporary map
        for layer in non_z_layers:
            temp_map.addLayer(layer)
        
        # Package the temporary map
        print("Packaging non-Z layers...")
        arcpy.management.PackageMap(
            in_map=temp_map,
            output_file=mpkx_path,
            convert_data="CONVERT",
            convert_arcsde_data="CONVERT_ARCSDE",
            extent=extent if extent else "#",
            apply_extent_to_arcsde="ALL" if extent else "NONE",
            select_related_rows="KEEP_ONLY_RELATED_ROWS",
            consolidate_to_one_fgdb="SINGLE_OUTPUT_WORKSPACE"
        )
        
        return mpkx_path
        
    finally:
        # Clean up temporary map
        aprx.deleteItem(temp_map)

def extract_7z(archive_path, extract_path, seven_zip_path):
    """
    Extracts a 7z archive using 7-Zip.
    
    Args:
        archive_path: Path to archive file
        extract_path: Directory to extract to
        seven_zip_path: Path to 7z.exe
        
    Raises:
        subprocess.CalledProcessError: If extraction fails
    """
    result = subprocess.run(
        [seven_zip_path, 'x', archive_path, f'-o{extract_path}', '-y'],
        capture_output=True,
        text=True
    )
    
    if result.returncode != 0:
        raise subprocess.CalledProcessError(
            result.returncode,
            result.args,
            result.stdout,
            result.stderr
        )

def find_geodatabase(search_path):
    """
    Finds a file geodatabase in the specified directory.
    
    Args:
        search_path: Directory to search in
        
    Returns:
        Path to the found geodatabase
        
    Raises:
        ValueError: If no geodatabase is found
    """
    for root, dirs, files in os.walk(search_path):
        for d in dirs:
            if d.endswith('.gdb'):
                return os.path.join(root, d)
    raise ValueError("No geodatabase found in package")

def preserve_metadata(map_obj, target_gdb):
    """
    Preserves metadata from map layers to the archived copies in the geodatabase.
    
    Args:
        map_obj: ArcGIS Pro Map object
        target_gdb: Path to the target geodatabase
    """
    for layer in map_obj.listLayers():
        if not layer.isGroupLayer:
            try:
                # Get the layer's metadata
                source_md = layer.metadata
                
                # Find the corresponding feature class in the archive geodatabase
                target_fc = os.path.join(target_gdb, arcpy.ValidateTableName(layer.name))
                
                if arcpy.Exists(target_fc):
                    # Get the target feature class's metadata object
                    target_md = md.Metadata(target_fc)
                    
                    if not target_md.isReadOnly:
                        # Copy metadata and save
                        target_md.copy(source_md)
                        target_md.save()
                        
                        print(f"Metadata preserved for: {layer.name}")
                    else:
                        print(f"Warning: Cannot write metadata for {layer.name} - target is read-only")
                else:
                    print(f"Warning: Could not find archived feature class for {layer.name}")
                    
            except Exception as e:
                print(f"Warning: Error processing metadata for {layer.name}: {str(e)}")

def add_extent_metadata(gdb_path, extent):
    """
    Adds extent information to the geodatabase's metadata.
    
    Args:
        gdb_path: Path to the geodatabase
        extent: Extent object containing boundary coordinates
    """
    try:
        gdb_md = md.Metadata(gdb_path)
        
        if not gdb_md.isReadOnly:
            extent_desc = (f"Data has been clipped to the following extent:\n"
                         f"XMin: {extent.XMin}\n"
                         f"YMin: {extent.YMin}\n"
                         f"XMax: {extent.XMax}\n"
                         f"YMax: {extent.YMax}")
            
            current_desc = gdb_md.description if gdb_md.description else ""
            gdb_md.description = current_desc + "\n\n" + extent_desc if current_desc else extent_desc
            
            gdb_md.save()
            print("Added extent information to geodatabase metadata")
    except Exception as e:
        print(f"Warning: Could not add extent metadata to geodatabase: {str(e)}")

def process_metadata(map_obj, gdb_path, extent=None):
    """
    Processes and preserves metadata for all layers in the geodatabase.
    
    Args:
        map_obj: ArcGIS Pro Map object
        gdb_path: Path to the target geodatabase
        extent: Optional extent information to add to metadata
    """
    print("Processing metadata...")
    preserve_metadata(map_obj, gdb_path)
    
    if extent:
        add_extent_metadata(gdb_path, extent)

def process_existing_gdb(map_obj, gdb_path, extent=None):
    """
    Processes metadata for an existing geodatabase extracted from an MPKX.
    
    Args:
        map_obj: ArcGIS Pro Map object
        gdb_path: Path to the existing geodatabase
        extent: Optional extent information to add to metadata
    """
    if not os.path.exists(gdb_path):
        raise ValueError(f"Geodatabase not found at: {gdb_path}")
        
    if not gdb_path.endswith('.gdb'):
        raise ValueError(f"Path does not point to a geodatabase: {gdb_path}")
    
    # Initialize a walk of the file geodatabase hierarchy
    walk = arcpy.da.Walk(gdb_path)

    # PackageMap prepends L0, L1, etc. to the file name; check with a regex and
    # rename if found
    for dirpath, dirnames, filenames in walk:
        for i, filename in enumerate(filenames):
            if re.match('L[0-9]', filename):
                prev_filename = filename
                # Remove first two characters from filename
                filename = filename[2:]
                # Names must start with an alpha character; advance down the
                # string until the first character is alphabetical
                while not filename[0].isalpha():
                    if len(filename) == 1:
                        # Layer name was only numbers/special characters; give
                        # a new valid name
                        filename = "No_Alpha_Characters_Layer_{i}"
                    else:
                        filename = filename[1:]
                print(f"Renaming {prev_filename} to {filename}")
                arcpy.management.Rename(os.path.join(dirpath, prev_filename),
                    os.path.join(dirpath, filename))
        
    process_metadata(map_obj, gdb_path, extent)
    print(f"Metadata processing completed for: {gdb_path}")

def identify_z_layers(map_obj, verbose = False):
    """
    Identifies layers with and without Z values in a map.
    
    Args:
        map_obj: ArcGIS Pro Map object
        
    Returns:
        Tuple of (non_z_layers, z_layers) lists
    """
    non_z_layers = []
    z_layers = []
    
    for layer in map_obj.listLayers():
        if not layer.isGroupLayer:
            try:
                desc = arcpy.Describe(layer)
                if hasattr(desc, 'hasZ') and desc.hasZ:
                    if verbose:
                        print(f"Found Z-enabled layer: {layer.name}")
                    z_layers.append(layer)
                else:
                    non_z_layers.append(layer)
            except Exception as e:
                print(f"Warning: identify_z_layers() could not process {layer.name}: {str(e)}")
                continue
    
    return non_z_layers, z_layers

def extract_existing_mpkx(mpkx_path, output_folder, map_obj=None, extent=None, z_layers=None, cleanup_temp=True):
    """
    Extracts an existing MPKX file to a geodatabase and optionally adds Z-enabled layers.
    
    Args:
        mpkx_path: Path to existing MPKX file
        output_folder: Directory where the geodatabase should be created
        map_obj: Optional Map object for layer name validation
        z_layers: Optional list of Z-enabled layers to add to the extracted geodatabase
        cleanup_temp: Whether to remove temporary files after completion
        
    Returns:
        Path to the extracted geodatabase
    """
    if not os.path.exists(mpkx_path):
        raise ValueError(f"MPKX file not found at: {mpkx_path}")
    
    if map_obj:
        map_name = arcpy.ValidateTableName(map_obj.name)
        map_crs = map_obj.spatialReference
    else:
        map_name = "map"
        map_crs = arcpy.SpatialReference(26916) # NAD83 UTM Zone 16N
    output_gdb = os.path.join(output_folder, f"{map_name}.gdb")
    temp_7z = os.path.join(output_folder, "temp.7z")
    
    # Find 7-Zip
    seven_zip_path = find_7zip_path()
    
    try:
        # Copy and rename MPKX to 7z
        print("Preparing package for extraction...")
        shutil.copy2(mpkx_path, temp_7z)
        
        # Extract using 7-Zip
        print("Extracting package...")
        extract_7z(temp_7z, output_folder, seven_zip_path)
        
        # Find and relocate the geodatabase
        source_gdb = find_geodatabase(output_folder)
        shutil.move(source_gdb, output_gdb)
        
        # Add Z-enabled layers if provided
        if z_layers:
            print("\nProcessing Z-enabled layers...")
            # Enable output overwrite
            arcpy.env.overwriteOutput = True
            for layer in z_layers:
                try:
                    feature_name = arcpy.ValidateTableName(layer.name)
                    out_fc = os.path.join(output_gdb, feature_name)
                    print(f"Copying and reprojecting {layer.name} to {out_fc}...", end = ' ')
                    # Copy features to local geodatabase
                    arcpy.management.CopyFeatures(layer, out_fc)
                    print("Copying complete...", end = ' ')
                    # Reproject to memory workspace
                    arcpy.management.Project(
                        out_fc,
                        r"memory\tempCopy",
                        out_coor_system=map_crs,
                        preserve_shape = "PRESERVE_SHAPE")
                    print("Reproject complete...", end = ' ')
                    # Overwrite with the reprojected feature
                    arcpy.management.CopyFeatures(r"memory\tempCopy", out_fc)
                    # Clear the in-memory workspace
                    arcpy.management.Delete(r"memory\tempCopy")
                    print("Done!")
                except Exception as e:
                    print(f"Warning: Failed to reproject {layer.name}: {str(e)}")
        
        # Cleanup if requested
        if cleanup_temp:
            if os.path.exists(mpkx_path):
                os.remove(mpkx_path)
            if os.path.exists(temp_7z):
                os.remove(temp_7z)
            print("Temporary files cleaned up")
        else:
            print(f"Temporary files preserved in: {output_folder}")
        
        print(f"Archive extracted successfully to: {output_gdb}")
        return output_gdb
        
    except Exception as e:
        print(f"Error: {str(e)}")
        if cleanup_temp and os.path.exists(output_folder):
            shutil.rmtree(output_folder)
        raise

## Example usage/main script run
if __name__ == "__main__":
    # Project file (input) and archive location (output)
    project_path = r"C:\Your\Project\path\here.aprx"
    output_location = r"C:\Your\output\path\here"
    # Name of the target map in the ArcGIS Pro project (accepts * as a wildcard)
    map_name = "Your map name here*"
    # Name of the layer in the map being used as the extent (bounding box)
    extent_layer_name = "Your extent layer here"
    
    try:
        # Open the project
        aprx = arcpy.mp.ArcGISProject(project_path)
        # Retrieve the map and extent
        map_obj = aprx.listMaps(map_name)[0]
        extent = get_layer_extent(map_obj, extent_layer_name)
        # Set the extent of the environment to the extent layer
        arcpy.env.extent = extent
        # Disable Z and M values
        arcpy.env.outputZFlag = "Disabled"
        arcpy.env.outputMFlag = "Disabled"
        
        # Create timestamped output folder
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        sanitized_name = arcpy.ValidateTableName(map_obj.name)
        output_folder = os.path.join(
            output_location,
            f"archive_{sanitized_name}_{timestamp}")
        os.makedirs(output_folder, exist_ok=True)
        
        # Identify Z and non-Z layers before packaging
        non_z_layers, z_layers = identify_z_layers(map_obj, verbose = True)
        
        # Create the MPKX with non-Z layers
        mpkx_path = package_map(aprx, map_obj, output_folder, extent=extent)
        
        # Extract the MPKX and add Z layers
        extracted_gdb = extract_existing_mpkx(
            mpkx_path,
            output_folder,
            map_obj,
            extent,
            z_layers,
            cleanup_temp=False)
        
        # Process metadata for the extracted geodatabase
        process_existing_gdb(map_obj, extracted_gdb)
        
        # Close the project connection
        del aprx
        
    except Exception as e:
        print(f"Error: {str(e)}")
        raise
