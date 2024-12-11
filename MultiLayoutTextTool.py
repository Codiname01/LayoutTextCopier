from pyautocad import Autocad, APoint
import sys
import time

# Conexión con AutoCAD
acad = Autocad()

# Mensaje para el usuario
acad.prompt("Por favor, selecciona un objeto de tipo texto (por ejemplo, una anotación) que deseas copiar en todos los layouts. Asegúrate de que sea un texto individual o texto multilinea visible.\n")

try:
    # Elimina el conjunto de selección existente si ya existe
    for ss in acad.doc.SelectionSets:
        if ss.Name == "TextoSeleccionado":
            ss.Delete()
            break

    # Seleccionar el texto manualmente
    seleccion = acad.doc.SelectionSets.Add("TextoSeleccionado")
    seleccion.SelectOnScreen()

    # Verifica si hay objetos seleccionados
    if hasattr(seleccion, 'Count') and seleccion.Count > 0:
        texto_original = seleccion[0]
        print(f"Se ha seleccionado un objeto de tipo: {texto_original.ObjectName}")

        if texto_original.ObjectName == "AcDbText" and hasattr(texto_original, "TextString"):
            # Texto simple (DBText)
            texto_contenido = texto_original.TextString
            texto_posicion = APoint(texto_original.InsertionPoint[0], texto_original.InsertionPoint[1])
            text_height = texto_original.Height
            es_mtext = False
        elif texto_original.ObjectName == "AcDbMText":
            # Texto multilinea (MText)
            # Obtener el contenido de texto
            if hasattr(texto_original, "TextString"):
                texto_contenido = texto_original.TextString
            else:
                print("No se encontró 'TextString' en el MText. Propiedades disponibles:")
                print(dir(texto_original))
                seleccion.Delete()
                sys.exit()

            # Obtener el punto de inserción del MText
            if hasattr(texto_original, "InsertionPoint"):
                texto_posicion = APoint(texto_original.InsertionPoint[0], texto_original.InsertionPoint[1])
            else:
                print("No se encontró la propiedad InsertionPoint en MText.")
                seleccion.Delete()
                sys.exit()

            # Obtener la altura del texto
            if hasattr(texto_original, "Height"):
                text_height = texto_original.Height
            elif hasattr(texto_original, "TextHeight"):
                text_height = texto_original.TextHeight
            else:
                print("No se encontró una propiedad de altura conocida (Height o TextHeight) en MText. Propiedades disponibles:")
                print(dir(texto_original))
                seleccion.Delete()
                sys.exit()

            es_mtext = True
        else:
            print("El objeto seleccionado no es texto válido (ni texto simple ni multilinea).")
            seleccion.Delete()
            sys.exit()

        # Recorremos todos los layouts del archivo sin cambiar ActiveLayout
        for layout in acad.doc.Layouts:
            if layout.Name != "Model":  # Ignoramos el espacio modelo
                print(f"Copiando texto en layout: {layout.Name}")
                block_layout = layout.Block

                # Insertamos el texto en el bloque del layout actual
                if es_mtext:
                    # Insertar como MText
                    ancho_mtext = 100  # Ajusta el ancho según tus necesidades
                    nuevo_mtext = block_layout.AddMText(texto_posicion, ancho_mtext, texto_contenido)
                    nuevo_mtext.Height = text_height
                else:
                    # Insertar como texto simple
                    block_layout.AddText(
                        texto_contenido,
                        texto_posicion,
                        text_height
                    )

                # Pausa breve para evitar saturar la comunicación COM
                time.sleep(0.5)

        print("Texto copiado exitosamente en todos los layouts.")

    else:
        print("No se seleccionó ningún objeto válido.")

    # Limpia la selección para evitar problemas futuros
    seleccion.Delete()

except Exception as e:
    error_code = getattr(e, 'errno', 'No disponible')
    print(f"Ocurrió un error: {e}\nCódigo de error: {error_code}")
    print("Sugerencias: 1. Verifique que AutoCAD esté abierto. 2. Asegúrese de que el texto seleccionado sea válido. 3. Compruebe que los layouts sean accesibles y editables.")
    # Eliminar selección si hubo un error
    if 'seleccion' in locals():
        seleccion.Delete()
    sys.exit()
