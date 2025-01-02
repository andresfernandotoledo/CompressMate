    function initializePage() {
            // Inicializar la página en el estado de "Comprimir"
            document.getElementById('action_type').value = 'compress';
            handleActionChange(); // Mostrar las opciones de compresión
            toggleCalculatorVisibility(); // Configurar la visibilidad inicial de la calculador
        }

        function handleActionChange() {
            const action = document.getElementById('action_type').value;
            // Ocultar todas las opciones primero
            document.getElementById('compress_options').classList.add('hidden');
            document.getElementById('convert_options').classList.add('hidden');

            // Mostrar la sección correspondiente
            if (action === 'compress') {
                document.getElementById('compress_options').classList.remove('hidden');
                toggleCompressOptions(); // Actualizar opciones de compresión basadas en el tipo de archivo por defecto
            } else if (action === 'convert') {
                document.getElementById('convert_options').classList.remove('hidden');
            }

            // Mostrar u ocultar la calculadora dependiendo de la acción seleccionada
             toggleCalculatorVisibility();
        }

        function toggleCalculatorVisibility() {
            const calculator = document.getElementById("calculator");
            const calculatorIcon = document.getElementById("calculator-icon");

            if (calculator.classList.contains("hidden")) {
                calculator.classList.remove("hidden");
                calculatorIcon.classList.add("hidden");
            } else {
                calculator.classList.add("hidden");
                calculatorIcon.classList.remove("hidden");
            }
          }


         // Opcional: si quieres que la calculadora se oculte cuando se haga clic fuera de ella
          document.addEventListener("click", function(event) {
          const calculator = document.getElementById("calculator");
          const calculatorIcon = document.getElementById("calculator-icon");
    
          if (!calculator.contains(event.target) && !calculatorIcon.contains(event.target)) {
               calculator.classList.add("hidden");
               calculatorIcon.classList.remove("hidden");
          }
        });

        function toggleCompressOptions() {
            const fileType = document.getElementById('file_type').value;
            const compressOptions = {
                imagen: 'image_quality',
                pdf: 'compression_quality',
                docx: 'compression_quality',
                pptx: 'compression_quality',
                xlsx: 'compression_quality',
                video: 'bitrate'
            };

            // Ocultar todas las opciones de compresión
            ['image_quality', 'compression_quality', 'bitrate'].forEach(function(id) {
                document.getElementById(id).classList.add('hidden');
            });

            // Mostrar la opción correspondiente
            const showDivId = compressOptions[fileType];
            if (showDivId) {
                document.getElementById(showDivId).classList.remove('hidden');
            }
        }

        function updateOutputFormats() {
            const fileType = document.getElementById('file_type_convert').value;
            const formatSelect = document.getElementById('output_format');
            
            // Limpiar opciones existentes
            formatSelect.innerHTML = '';
            
            // Opciones de formatos según el tipo de archivo
            const formats = {
                'imagen': ['jpeg', 'png', 'bmp', 'gif', 'tiff'],
                'pdf': ['docx', 'pptx', 'xlsx', 'jpeg', 'png'],
                'docx': ['pdf', 'pptx', 'xlsx'],
                'pptx': ['pdf', 'docx', 'xlsx'],
                'xlsx': ['pdf', 'docx', 'pptx'],
                'video': ['mp4', 'mov', 'wmv', 'avi', 'flv', 'webm']
            };
            
            if (formats[fileType]) {
                formats[fileType].forEach(format => {
                    const option = document.createElement('option');
                    option.value = format;
                    option.textContent = format.toUpperCase(); // Mostrar en mayúsculas
                    formatSelect.appendChild(option);
                });
            }
        }


        function updateFileName() {
            const fileInput = document.getElementById('file');
            const fileNameSpan = document.getElementById('file-name');

             if (fileInput.files.length > 0) {
                 fileNameSpan.textContent = fileInput.files[0].name;
             } else {
                 fileNameSpan.textContent = 'Seleccionar archivo';
             }
           }

        document.getElementById('upload-form').addEventListener('submit', async function(event) {
            event.preventDefault(); // Evitar el envío tradicional del formulario

            showLoader(); // Mostrar el spinner

            // Crear un FormData para enviar el formulario de manera asincrónica
            const formData = new FormData(this);

            try {
                // Enviar el formulario usando fetch
                const response = await fetch(this.action, {
                    method: this.method,
                    body: formData
                });

                if (!response.ok) {
                    throw new Error('Error en la respuesta del servidor');
                }

                // Aquí puedes manejar la respuesta del servidor si es necesario
                const result = await response.json(); // Suponiendo que el servidor devuelve JSON
                alert(result.message || 'Documento procesado exitosamente');

            } catch (error) {
                // Manejar errores aquí
                alert('Error al procesar el documento: ' + error.message);

            } finally {
                hideLoader(); // Restaurar el botón a su estado original
            }
        });

        function showLoader() {
            document.getElementById('button-text').style.display = 'none';
            document.getElementById('spinner').classList.remove('hidden');
        }

        function hideLoader() {
            document.getElementById('button-text').style.display = 'inline-block';
            document.getElementById('spinner').classList.add('hidden');
        }
         
        function convertBytes() {
            var inputValue = parseFloat(document.getElementById('input_value').value);
            var unit = document.getElementById('input_unit').value;
            var result = document.getElementById('conversion_result');
            
            if (isNaN(inputValue)) {
                result.textContent = 'Por favor, ingresa un valor válido.';
                return;
            }

            var bytes;
            switch (unit) {
                case 'bytes':
                    bytes = inputValue;
                    break;
                case 'kb':
                    bytes = inputValue * 1024;
                    break;
                case 'mb':
                    bytes = inputValue * 1024 * 1024;
                    break;
                case 'gb':
                    bytes = inputValue * 1024 * 1024 * 1024;
                    break;
                default:
                    bytes = inputValue;
                    break;
            }

            var kb = bytes / 1024;
            var mb = kb / 1024;
            var gb = mb / 1024;

            result.innerHTML = `Bytes: ${bytes.toFixed(2)} B<br>KB: ${kb.toFixed(2)} KB<br>MB: ${mb.toFixed(2)} MB<br>GB: ${gb.toFixed(2)} GB`;
        }