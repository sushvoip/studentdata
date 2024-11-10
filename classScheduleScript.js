document.addEventListener('DOMContentLoaded', function() {
    const filePath = './studentdata.xlsx'; // Relative path to the Excel file

    fetch(filePath)
        .then(response => response.arrayBuffer())
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });

            const classScheduleSheet = XLSX.utils.sheet_to_json(workbook.Sheets['ClassSchedule']);

            // Format time to 'hh:mm a.m./p.m.'
            function formatTime(timeValue) {
                if (typeof timeValue === 'number') {
                    const totalMinutes = Math.round(timeValue * 1440);
                    const hours = Math.floor(totalMinutes / 60);
                    const minutes = totalMinutes % 60;
                    const date = new Date(1970, 0, 1, hours, minutes);
                    return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: true, timeZone: 'America/New_York' });
                } else if (typeof timeValue === 'string') {
                    try {
                        let [time, period] = timeValue.trim().toUpperCase().split(' ');
                        let [hours, minutes] = time.split(':');
                        hours = parseInt(hours, 10);
                        minutes = parseInt(minutes, 10);
                        if (period === 'PM' && hours < 12) hours += 12;
                        if (period === 'AM' && hours === 12) hours = 0;
                        return new Date(1970, 0, 1, hours, minutes).toLocaleTimeString('en-US', {
                            hour: '2-digit',
                            minute: '2-digit',
                            hour12: true,
                            timeZone: 'America/New_York'
                        });
                    } catch (e) {
                        console.error('Invalid time value:', timeValue);
                        return '';
                    }
                }
                console.error('Unknown time value type:', timeValue);
                return '';
            }

            // Populate the ClassSchedule tab
            const classScheduleTableBody = document.getElementById('ClassScheduleTable').getElementsByTagName('tbody')[0];
            const classScheduleHeaders = Object.keys(classScheduleSheet[0]);

            // Create the header row
            const headerRow = document.createElement('tr');
            classScheduleHeaders.forEach(header => {
                const cell = document.createElement('th');
                cell.textContent = header;
                headerRow.appendChild(cell);
            });
            classScheduleTableBody.appendChild(headerRow);

            // Create the body rows
            classScheduleSheet.forEach(classInfo => {
                const row = document.createElement('tr');
                classScheduleHeaders.forEach(key => {
                    const cell = document.createElement('td');
                    if (key === 'Start Time' || key === 'End Time') {
                        cell.textContent = formatTime(classInfo[key]);
                    } else {
                        cell.textContent = classInfo[key] !== undefined ? classInfo[key] : '';
                    }
                    row.appendChild(cell);
                });
                classScheduleTableBody.appendChild(row);
            });

        })
        .catch(error => console.error('Error fetching or parsing file:', error));
});
