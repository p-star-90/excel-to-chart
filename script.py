import js
from pyodide.ffi import create_proxy
from js import document, File, URL, Image, console
import io
import base64
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # Use Agg backend for non-interactive plotting

# Global variables to store the data
excel_data = None
file_name = None

# Set up event listeners when the page loads
def setup():
    file_input = document.getElementById("excel-file")
    file_input.addEventListener("change", create_proxy(handle_file_select))
    
    generate_btn = document.getElementById("generate-btn")
    generate_btn.addEventListener("click", create_proxy(generate_chart))
    
    # Add event listener for chart type change
    radio_buttons = document.getElementsByName("chart-type")
    for button in radio_buttons:
        button.addEventListener("change", create_proxy(toggle_chart_options))
    
    # Initial toggle of chart options
    toggle_chart_options()
    
    # Add debug message
    status_message = document.getElementById("status-message")
    status_message.innerHTML = "Application initialisée. Veuillez télécharger un fichier Excel."
    status_message.className = "info"

def toggle_chart_options(event=None):
    chart_type = "bar"  # Default
    radio_buttons = document.getElementsByName("chart-type")
    for button in radio_buttons:
        if button.checked:
            chart_type = button.value
            break
    
    bar_options = document.getElementById("bar-chart-options")
    pie_options = document.getElementById("pie-chart-options")
    
    if chart_type == "bar":
        bar_options.style.display = "block"
        pie_options.style.display = "none"
    else:
        bar_options.style.display = "none"
        pie_options.style.display = "block"

def handle_file_select(event):
    global excel_data, file_name
    
    file_info = document.getElementById("file-info")
    status_message = document.getElementById("status-message")
    status_message.innerHTML = ""
    status_message.className = ""
    
    try:
        if event.target.files.length > 0:
            file = event.target.files.item(0)
            file_name = file.name
            
            console.log(f"File selected: {file_name}")
            file_info.innerText = f"Fichier sélectionné: {file_name}"
            
            excel_data = None
            
            x_column = document.getElementById("x-column")
            y_column = document.getElementById("y-column")
            x_column.innerHTML = "<option value=''>Chargement...</option>"
            y_column.innerHTML = "<option value=''>Chargement...</option>"
            x_column.disabled = True
            y_column.disabled = True
            
            status_message.innerHTML = "Chargement du fichier..."
            status_message.className = "loading"
            
            read_file(file)
        else:
            file_info.innerText = "Aucun fichier sélectionné"
            excel_data = None
    except Exception as e:
        console.error(f"Error in handle_file_select: {str(e)}")
        status_message.innerHTML = f"Erreur lors de la sélection du fichier: {str(e)}"
        status_message.className = "error"

def populate_column_selectors():
    global excel_data
    
    if excel_data is None or excel_data.empty:
        return
    
    x_column = document.getElementById("x-column")
    y_column = document.getElementById("y-column")
    
    x_column.innerHTML = ""
    y_column.innerHTML = ""
    
    for i, col in enumerate(excel_data.columns):
        x_option = document.createElement("option")
        x_option.value = str(i)
        x_option.text = col
        x_column.appendChild(x_option)
        
        y_option = document.createElement("option")
        y_option.value = str(i)
        y_option.text = col
        y_column.appendChild(y_option)
    
    x_column.value = "0"
    if len(excel_data.columns) > 1:
        y_column.value = "1"
    
    x_column.disabled = False
    y_column.disabled = False

def read_file(file):
    global excel_data
    
    def process_file(event):
        global excel_data
        status_message = document.getElementById("status-message")
        
        try:
            array_buffer = event.target.result
            console.log("File loaded into ArrayBuffer")
            
            bytes_data = bytes(js.Uint8Array.new(array_buffer))
            console.log(f"Bytes data length: {len(bytes_data)}")
            
            try:
                excel_data = pd.read_excel(io.BytesIO(bytes_data), engine='openpyxl')
                console.log("File read with openpyxl engine")
            except Exception as e1:
                console.error(f"openpyxl engine failed: {str(e1)}")
                try:
                    excel_data = pd.read_excel(io.BytesIO(bytes_data), engine='xlrd')
                    console.log("File read with xlrd engine")
                except Exception as e2:
                    console.error(f"xlrd engine failed: {str(e2)}")
                    raise Exception(f"Failed to read Excel file with both engines: {str(e1)} and {str(e2)}")
            
            if excel_data is not None and not excel_data.empty:
                console.log(f"Data loaded successfully. Shape: {excel_data.shape}")
                console.log(f"Columns: {list(excel_data.columns)}")
                
                populate_column_selectors()
                
                status_message.innerHTML = f"Fichier chargé avec succès! Lignes: {excel_data.shape[0]}, Colonnes: {excel_data.shape[1]}"
                status_message.className = "success"
            else:
                raise Exception("Le fichier Excel est vide ou n'a pas pu être lu correctement")
            
        except Exception as e:
            console.error(f"Error in process_file: {str(e)}")
            status_message.innerHTML = f"Erreur lors du chargement du fichier: {str(e)}"
            status_message.className = "error"
            excel_data = None
    
    try:
        reader = js.FileReader.new()
        reader.onload = create_proxy(process_file)
        reader.onerror = create_proxy(lambda e: console.error(f"FileReader error: {e}"))
        reader.readAsArrayBuffer(file)
    except Exception as e:
        console.error(f"Error setting up FileReader: {str(e)}")
        status_message = document.getElementById("status-message")
        status_message.innerHTML = f"Erreur lors de la configuration du lecteur de fichier: {str(e)}"
        status_message.className = "error"

def optimize_pie_chart(labels, values, min_percent=2.0):
    """Optimise les données du graphique circulaire pour une meilleure lisibilité."""
    total = sum(values)
    percentages = [v/total * 100 for v in values]
    
    # Séparer les données en fonction du seuil
    main_data = [(l, v, p) for l, v, p in zip(labels, values, percentages) if p >= min_percent]
    small_data = [(l, v, p) for l, v, p in zip(labels, values, percentages) if p < min_percent]
    
    # Si des données sont regroupées, les combiner
    if small_data:
        other_value = sum(v for _, v, _ in small_data)
        other_percent = sum(p for _, _, p in small_data)
        main_data.append(("Autres", other_value, other_percent))
        
        # Créer une note pour les petites valeurs
        small_labels = [f"{l} ({p:.1f}%)" for l, _, p in small_data]
    else:
        small_labels = []
    
    # Trier les données principales par valeur décroissante
    main_data.sort(key=lambda x: x[1], reverse=True)
    
    return (
        [x[0] for x in main_data],  # labels
        [x[1] for x in main_data],  # values
        small_labels  # liste des petites valeurs pour la légende
    )

def generate_chart(event=None):
    global excel_data, file_name
    
    status_message = document.getElementById("status-message")
    chart_display = document.getElementById("chart-display")
    
    console.log(f"Generate chart called. excel_data is {'not None' if excel_data is not None else 'None'}")
    
    if excel_data is None:
        status_message.innerHTML = "Veuillez d'abord télécharger un fichier Excel"
        status_message.className = "error"
        return
    
    # Get chart type and options
    chart_type = None
    radio_buttons = document.getElementsByName("chart-type")
    for button in radio_buttons:
        if button.checked:
            chart_type = button.value
            break
    
    try:
        x_column_idx = int(document.getElementById("x-column").value)
        y_column_idx = int(document.getElementById("y-column").value)
    except:
        status_message.innerHTML = "Veuillez sélectionner des colonnes valides"
        status_message.className = "error"
        return
    
    # Get customization options
    chart_title = document.getElementById("chart-title").value
    x_axis_title = document.getElementById("x-axis-title").value
    y_axis_title = document.getElementById("y-axis-title").value
    
    # Get pie chart specific options
    min_percent = float(document.getElementById("min-percent").value)
    show_legend = document.getElementById("show-legend").checked
    show_percentages = document.getElementById("show-percentages").checked
    
    # Get additional options
    use_default_titles = document.getElementById("use-default-titles").checked
    
    status_message.innerHTML = "Génération du graphique..."
    status_message.className = "loading"
    
    try:
        chart_display.innerHTML = ""
        plt.figure(figsize=(12, 8))  # Augmenter la taille du graphique
        
        if excel_data.shape[1] < 2:
            raise ValueError("Les données doivent avoir au moins deux colonnes (étiquettes et valeurs)")
        
        labels = excel_data.iloc[:, x_column_idx].tolist()
        values = excel_data.iloc[:, y_column_idx].tolist()
        
        if chart_type == "bar":
            plt.bar(labels, values, color='#3498db')
            
            # Définir les titres des axes uniquement si spécifiés ou si l'option par défaut est activée
            if x_axis_title:
                plt.xlabel(x_axis_title)
            elif use_default_titles:
                plt.xlabel(excel_data.columns[x_column_idx])
                
            if y_axis_title:
                plt.ylabel(y_axis_title)
            elif use_default_titles:
                plt.ylabel(excel_data.columns[y_column_idx])
                
            plt.xticks(rotation=45, ha='right')
            
        elif chart_type == "pie":
            # Optimiser les données pour le graphique circulaire
            opt_labels, opt_values, small_labels = optimize_pie_chart(labels, values, min_percent)
            
            # Créer le graphique circulaire
            patches, texts, autotexts = plt.pie(
                opt_values,
                labels=opt_labels if not show_legend else None,
                autopct='%1.1f%%' if show_percentages else None,
                startangle=90,
                counterclock=False,  # Sens horaire
                pctdistance=0.85,    # Position des pourcentages
                labeldistance=1.1    # Position des étiquettes
            )
            
            # Ajuster les propriétés du texte
            if show_percentages:
                plt.setp(autotexts, size=8, weight="bold")
            if not show_legend:
                plt.setp(texts, size=10)
            
            # Ajouter une légende si demandé
            if show_legend:
                plt.legend(
                    patches,
                    opt_labels,
                    title="Légende",
                    loc="center left",
                    bbox_to_anchor=(1, 0, 0.5, 1)
                )
            
            # Ajouter la note des petites valeurs si nécessaire
            if small_labels:
                note = "Valeurs < {}%:\n{}".format(
                    min_percent,
                    "\n".join(small_labels)
                )
                plt.figtext(
                    0.95, 0.02,
                    note,
                    ha='right',
                    va='bottom',
                    fontsize=8,
                    style='italic'
                )
            
            plt.axis('equal')
        
        # Set chart title uniquement si spécifié ou si l'option par défaut est activée
        if chart_title:
            plt.title(chart_title, pad=20)
        elif use_default_titles:
            if chart_type == "bar":
                plt.title(f"Graphique à barres - {file_name}", pad=20)
            elif chart_type == "pie":
                plt.title(f"Graphique circulaire - {file_name}", pad=20)
        
        # Ajuster la mise en page
        if chart_type == "pie" and show_legend:
            plt.tight_layout(rect=[0, 0.1, 0.85, 0.9])  # Laisser de l'espace pour la légende
        else:
            plt.tight_layout(rect=[0, 0.1, 1, 0.9])
        
        # Save and display
        buf = io.BytesIO()
        plt.savefig(buf, format='png', dpi=100, bbox_inches='tight')
        buf.seek(0)
        
        img_str = base64.b64encode(buf.read()).decode('utf-8')
        
        img_element = document.createElement("img")
        img_element.src = f"data:image/png;base64,{img_str}"
        img_element.style.maxWidth = "100%"
        img_element.style.height = "auto"
        chart_display.appendChild(img_element)
        
        status_message.innerHTML = "Graphique généré avec succès!"
        status_message.className = "success"
        
        plt.close()
        
    except Exception as e:
        console.error(f"Error generating chart: {str(e)}")
        status_message.innerHTML = f"Erreur lors de la génération du graphique: {str(e)}"
        status_message.className = "error"
        plt.close()

# Initialize the page
setup()