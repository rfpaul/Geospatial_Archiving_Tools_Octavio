import arcpy
import os
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

def create_layer_summary(project_path, map_name, output_folder):
    """
    Creates a human-readable summary of all layers in a specified map and exports their metadata.

    Args:
        project_path (str): Path to the ArcGIS Pro project file (.aprx)
        map_name (str): Name of the map to analyze
        output_folder (str): Base folder for outputs

    Returns:
        tuple: (Path to summary file, Path to metadata folder)
    """
    # Setup output folders
    output_folder = Path(output_folder)
    metadata_folder = output_folder / "metadata"
    os.makedirs(metadata_folder, exist_ok=True)

    try:
        # Open the project and get the map
        aprx = arcpy.mp.ArcGISProject(project_path)
        target_map = aprx.listMaps(map_name)[0]
        map_name = target_map.name
        summary_file = output_folder / f"{map_name}_layer_summary.txt"
        if not target_map:
            raise ValueError(f"Map matching '{map_name}' not found in project")

        # Create the summary file
        with open(summary_file, 'w', encoding='utf-8') as f:
            f.write(f"Layer Summary for Map: {map_name}\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 80 + "\n\n")

            # Process each layer
            layers = target_map.listLayers()
            for layer in layers:
                if layer.isGroupLayer:
                    continue

                # Get basic layer info
                f.write(f"Layer Name: {layer.name}\n")

                # Get source information
                try:
                    if hasattr(layer, 'dataSource'):
                        f.write(f"Source Location: {layer.dataSource}\n")
                    else:
                        f.write("Source Location: \n")
                except:
                    f.write("Source Location: \n")

                # Get spatial reference information
                try:
                    desc = arcpy.Describe(layer)
                    if hasattr(desc, 'spatialReference'):
                        sr = desc.spatialReference
                        f.write(f"CRS/Projection: {sr.name} ({sr.factoryCode})\n")
                    else:
                        f.write("CRS/Projection: \n")
                except:
                    f.write("CRS/Projection: \n")

                # Get metadata information
                try:
                    md = layer.metadata
                    if md:
                        # Export metadata to XML
                        xml_filename = f"{layer.name.replace(' ', '_')}_metadata.xml"
                        xml_path = metadata_folder / xml_filename
                        md.saveAsXML(str(xml_path))

                        # Extract metadata fields
                        f.write(f"Author: {md.credits if md.credits else ''}\n")
                        # Export metadata to get dates from XML
                        xml_path_str = str(xml_path)
                        md.saveAsXML(xml_path_str)

                        # Parse XML to get creation and modification dates
                        creation_date = ""
                        modification_date = ""
                        try:
                            tree = ET.parse(xml_path_str)
                            root = tree.getroot()

                            # Look for dates in different possible locations
                            # Try CreaDate and ModDate first
                            crea_date = root.find(".//CreaDate")
                            mod_date = root.find(".//ModDate")

                            # If not found, try alternative metadata date fields
                            if crea_date is None:
                                crea_date = root.find(".//{http://www.isotc211.org/2005/gmd}dateStamp")
                            if mod_date is None:
                                mod_date = root.find(".//{http://www.isotc211.org/2005/gmd}dateStamp")

                            creation_date = crea_date.text if crea_date is not None else ""
                            modification_date = mod_date.text if mod_date is not None else ""

                        except Exception as e:
                            print(f"Warning: Could not parse dates from metadata XML for layer {layer.name}: {str(e)}")

                        f.write(f"Created: {creation_date}\n")
                        f.write(f"Last modified: {modification_date}\n")
                        # Convert to relative path for the summary
                        relative_path = os.path.relpath(xml_path, output_folder)
                        f.write(f"XML Metadata Location: {relative_path}\n")
                    else:
                        f.write("Author: \n")
                        f.write("Created: \n")
                        f.write("Last modified: \n")
                        f.write("XML Metadata Location: \n")
                except Exception as e:
                    print(f"Warning: Could not process metadata for layer {layer.name}: {str(e)}")
                    f.write("Author: \n")
                    f.write("Created: \n")
                    f.write("Last modified: \n")
                    f.write("XML Metadata Location: \n")

                f.write("\n" + "-" * 80 + "\n\n")

        return str(summary_file), str(metadata_folder)

    except Exception as e:
        raise Exception(f"Error processing map layers: {str(e)}")
    finally:
        if 'aprx' in locals():
            del aprx

def main():
    """
    Example usage of the create_layer_summary function.
    """
    # Example parameters
    project_path = r"C:\Your\Project\path\here.aprx"
    map_name = "Your map name here*"
    output_folder = r"C:\Your\output\path\here"

    try:
        summary_file, metadata_folder = create_layer_summary(
            project_path,
            map_name,
            output_folder
        )
        print(f"Summary file created: {summary_file}")
        print(f"Metadata exported to: {metadata_folder}")
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
