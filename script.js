document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('activityForm').addEventListener('submit', handleFormSubmit);

    function handleFormSubmit(event) {
        event.preventDefault(); // Evitar el envío automático del formulario

        const activity = document.getElementById('activity').value;
        const employee = document.getElementById('employee').value;
        const estado = document.getElementById('estado').value;
        const startDate = document.getElementById('startDate').value;
        const endDate = document.getElementById('endDate').value;
       

        if (!activity || !employee || !estado || !startDate || !endDate) {
            alert('Por favor, complete todos los campos.');
            return;
        }

        fetch('http://localhost:3000/save-activity', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                activity,
                employee,
                startDate,
                endDate,
                estado
                
            })
        })
        .then(response => {
            if (!response.ok) {
                throw new Error(`Error al registrar la actividad: ${response.statusText}`);
            }
            return response.text();
        })
        .then(message => {
            alert(message);

            const tableBody = document.getElementById('activityTableBody');
            const newRow = document.createElement('tr');

            newRow.innerHTML = `
                <td>${activity}</td>
                <td>${employee}</td>
                <td>${estado}</td>
                <td>${startDate}</td>
                <td>${endDate}</td>
               
            `;

            tableBody.appendChild(newRow);

            document.getElementById('activityForm').reset();
        })
        .catch(error => {
            console.error('Error al registrar la actividad:', error);
            alert('Hubo un error al registrar la actividad. Consulta la consola para más detalles.');
        });
    }
});
