import traceback
import adsk.core
import adsk.fusion
import os, sys

sys.path.append(os.path.join(os.path.dirname(__file__), "packages"))
from openpyxl import Workbook

app = adsk.core.Application.get()
ui = app.userInterface

def get_unique_filename(file_path):
    """
    Si el archivo ya existe, añade un número entre paréntesis al nombre del archivo.
    """
    base, extension = os.path.splitext(file_path)
    counter = 1
    new_file_path = file_path
    while os.path.exists(new_file_path):
        new_file_path = f"{base} ({counter}){extension}"
        counter += 1
    return new_file_path

def run(_context: str):
    try:        
        app.log("RUNNING SCRIPT...")
        
        # Obtener el documento activo
        design = app.activeProduct

        if not isinstance(design, adsk.fusion.Design):
            ui.messageBox('No active Fusion design')
            return

        # Preguntar la cantidad de muebles
        qty_input = ui.inputBox('Muebles:', 'Cantidad de Muebles', '1')

        if not qty_input[1]:  # Si el usuario no cancela
            # Obtener el componente raíz
            root_comp = design.rootComponent
            # Obtener la unidad activa del diseño (longitud)
            unit_symbol = design.fusionUnitsManager.defaultLengthUnits

            # Diccionario para contar ocurrencias de componentes
            component_counts = {}
            
            real_components_count = 0
            
            # Diccionario para almacenar los cuerpos con sus dimensiones, usando descripción como clave
            body_dimensions = {}

            # Función para calcular las dimensiones de un body
            def get_body_dimensions(body):
                target_unit = unit_symbol
                bounding_box = body.boundingBox
                dim_x_cm = abs(bounding_box.maxPoint.x - bounding_box.minPoint.x)
                dim_y_cm = abs(bounding_box.maxPoint.y - bounding_box.minPoint.y)
                dim_z_cm = abs(bounding_box.maxPoint.z - bounding_box.minPoint.z)

                dim_x = round(design.fusionUnitsManager.convert(dim_x_cm, 'cm', target_unit), 2)
                dim_y = round(design.fusionUnitsManager.convert(dim_y_cm, 'cm', target_unit), 2)
                dim_z = round(design.fusionUnitsManager.convert(dim_z_cm, 'cm', target_unit), 2)

                dimensions = [dim_x, dim_y, dim_z]
                dimensions.sort(reverse=True)

                return dimensions[0], dimensions[1], dimensions[2]

            # Procesar cuerpos en el componente raíz (solo visibles)
            for body in root_comp.bRepBodies:
                if body.isVisible:  # Verificar si el cuerpo es visible
                    x, y, z = get_body_dimensions(body)
                    description = f'{root_comp.name}-{body.name}'
                    # Contar como 1 para el root component inicialmente
                    component_counts[root_comp.name] = component_counts.get(root_comp.name, 0) + 1
                    body_dimensions[description] = [1, x, y, z, description]

            # Procesar cuerpos de las ocurrencias (solo visibles)
            all_occurrences = root_comp.allOccurrences
            app.log(f'occurrences: {len(all_occurrences)}')
            for occurrence in all_occurrences:
                if occurrence.isVisible:  # Verificar si la ocurrencia es visible
                    comp = occurrence.component
                    comp_name = comp.name
                    app.log(f' ========== analizing {comp_name} ========== ')
                    real_components_count += 1
                    # Contar ocurrencias del componente
                    component_counts[comp_name] = component_counts.get(comp_name, 0) + 1
                    qty = component_counts[comp_name]
                    
                    for body in comp.bRepBodies:
                        if body.isVisible:  # Verificar si el cuerpo es visible
                            x, y, z = get_body_dimensions(body)
                            description = f'{comp_name}-{body.name}'
                            if description in body_dimensions:
                                # Si el cuerpo ya existe, actualizar solo la cantidad
                                app.log(f' - incrementing {body.name} to {qty}')
                                body_dimensions[description][0] = qty
                            else:
                                # Si es nuevo, agregar con la cantidad actual
                                app.log(f' - adding {body.name}.')
                                body_dimensions[description] = [qty, x, y, z, description]
            
                else:
                    msg = f'ignoring {occurrence.component.name} due is not visible'
                    app.log(msg)
            app.log(' ============================= ')
            
        
            furniture_count = int(qty_input[0])
            app.log(f'Multipling QTDs for {furniture_count} Instances')

            # Multiplicar cada QTD por la cantidad de muebles
            for key in body_dimensions:
                original_value = body_dimensions[key][0]
                new_value = original_value * furniture_count
                body_dimensions[key][0] = new_value                
        
            # Crear un nuevo libro de Excel
            wb = Workbook()
            ws = wb.active
            ws.title = "Body Dimensions"

            # Escribir encabezados
            headers = [f'QTD.', f'L {{{unit_symbol}}}', f'A {{{unit_symbol}}}', f'E {{{unit_symbol}}}', 'Descripción']
            ws.append(headers)

            # Escribir los datos (convertir valores del diccionario a lista)
            for row in body_dimensions.values():
                ws.append(row)

            # Guardar el archivo en el escritorio
            desktop_path = os.path.join(os.path.expanduser("~"), "Downloads")
            file_path = os.path.join(desktop_path, f"{root_comp.name}.xlsx")
            unique_file_path = get_unique_filename(file_path)
            wb.save(unique_file_path)

            # Message:
            msg = (
                "Found:\n"
                f"• {real_components_count} Components\n"
                f"• {len(body_dimensions)} Unique Bodies\n\n"  
                f"Data saved to Excel file: {unique_file_path}\n"
            )
            
            app.log(msg)
            ui.messageBox(msg)
            app.log("END SCRIPT...")
        
        else:
            app.log('canceled.')

    except:
        app.log(f'Failed:\n{traceback.format_exc()}')
        ui.messageBox(f'Error occurred:\n{traceback.format_exc()}')