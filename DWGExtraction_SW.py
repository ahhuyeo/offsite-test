from pyautocad import Autocad, APoint
import pythoncom

def extract_layer_and_object_count(dwg_path, output_file):
    try:
        # Initialize COM interface
        pythoncom.CoInitialize()
        acad = Autocad(create_if_not_exists=True)
        
        # Open the DWG file
        acad.Application.Documents.Open(dwg_path)

        # Get the layers and their names
        layers = acad.doc.Layers
        layer_names = [layer.Name for layer in layers]
        
        # Initialize layer object count dictionary
        layer_object_counts = {layer_name: 0 for layer_name in layer_names}
        
        # Count the entities per layer
        modelspace = acad.doc.ModelSpace
        for entity in modelspace:
            if entity.Layer in layer_object_counts:
                layer_object_counts[entity.Layer] += 1

        # Total counts
        layer_count = len(layer_names)
        total_entities = sum(layer_object_counts.values())

        # Write results to the output file
        with open(output_file, 'w') as file:
            file.write(f"Layer count: {layer_count}\n")
            file.write(f"Object count: {total_entities}\n")
            file.write("Layers and their object counts:\n")
            for name in layer_names:
                count = layer_object_counts[name]
                if count > 0:
                    file.write(f"  {name}: {count}\n")

        print(f"Results written to {output_file}")

    except Exception as e:
        print(f"Error: {e}")

    finally:
        # Properly close AutoCAD instance
        try:
            acad.Application.Quit()
        except Exception as e:
            print(f"Error closing AutoCAD: {e}")

# Example usage
dwg_file_path = r'C:\Users\siakwee\OneDrive - Shin-Nippon Industries Sdn Bhd\Desktop\programming\PAM116CWD-RD9203.dwg'
output_text_file = r'C:\Users\siakwee\OneDrive - Shin-Nippon Industries Sdn Bhd\Desktop\programming\output.txt'

extract_layer_and_object_count(dwg_file_path, output_text_file)
