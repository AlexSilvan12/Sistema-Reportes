function mostrarCampoOtroCertificador() {
    const select = document.getElementById('certificador_select');
    const campoOtro = document.getElementById('campo_otro_certificador');
    campoOtro.style.display = select.value === 'otro' ? 'block' : 'none';
}

function actualizarNombreCliente() {
    const select = document.getElementById('cliente-select');
    const selectedOption = select.options[select.selectedIndex];
    const nombre = selectedOption.getAttribute('data-nombre');
    document.getElementById('nombre_cliente').value = nombre;
}

// Ejecutar al cargar la página también (por si hay uno preseleccionado)
document.addEventListener('DOMContentLoaded', actualizarNombreCliente);


function actualizarNombreEspecialista() {
  const select = document.getElementById('especialista-select');
  const selected = select.options[select.selectedIndex];
  const nombre = selected.getAttribute('data-nombre');
  document.getElementById('nombre_especialista').value = nombre;
}

// Ejecutar al cargar también
document.addEventListener('DOMContentLoaded', actualizarNombreEspecialista);


