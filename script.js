document.addEventListener('DOMContentLoaded', function() {
    const filePath = './studentdata.xlsx'; // Relative path to the Excel file

    fetch(filePath)
        .then(response => response.arrayBuffer())
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });

            const studentSheet = XLSX.utils.sheet_to_json(workbook.Sheets['Student']);
            const classScheduleSheet = XLSX.utils.sheet_to_json(workbook.Sheets['ClassSchedule']);

            // Filter active students
            const activeStudents = studentSheet.filter(student => student['Active'] === 'Y');

            // Format date to 'dd MMM yyyy'
            function formatDate(dateString) {
                if (!dateString) return '';
                if (typeof dateString === 'number') {
                    const excelDate = new Date((dateString - 25569) * 86400 * 1000);
                    return new Intl.DateTimeFormat('en-US', { day: '2-digit', month: 'short', year: 'numeric' }).format(excelDate);
                }
                try {
                    const date = new Date(dateString);
                    return new Intl.DateTimeFormat('en-US', { day: '2-digit', month: 'short', year: 'numeric' }).format(date);
                } catch (e) {
                    console.error('Invalid date value:', dateString);
                    return '';
                }
            }

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

            // Sort data by day and start time
            function sortDataByDayAndTime(data) {
                const daysOfWeek = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
                return data.sort((a, b) => {
                    const dayA = daysOfWeek.indexOf(a.Day);
                    const dayB = daysOfWeek.indexOf(b.Day);
                    if (dayA === dayB) {
                        const timeA = new Date(`1970/01/01 ${a['Start Time']}`);
                        const timeB = new Date(`1970/01/01 ${b['Start Time']}`);
                        return timeA - timeB;
                    }
                    return dayA - dayB;
                });
            }

            // Function to get a color based on index
            function getColor(index) {
                const colors = ['#FF9999', '#FFCC99', '#FFFF99', '#CCFF99', '#99FF99', '#99FFCC', '#99FFFF', '#99CCFF', '#9999FF', '#CC99FF'];
                return colors[index % colors.length];
            }

            try {
                // Group data by day and start time
                const groupedData = classScheduleSheet.reduce((acc, classInfo) => {
                    const students = activeStudents.filter(student => student['Class ID'] === classInfo['Class ID']);
                    students.forEach(student => {
                        const mergedItem = {
                            ...student,
                            ...classInfo,
                            'Date Enrolled': formatDate(student['Date Enrolled']),
                            'Start Time': formatTime(classInfo['Start Time']),
                            'End Time': formatTime(classInfo['End Time']),
                            'Online': classInfo['Online']
                        };
                        const key = `${mergedItem.Day}-${mergedItem['Start Time']}`;
                        if (!acc[key]) {
                            acc[key] = [];
                        }
                        acc[key].push(mergedItem);
                    });
                    return acc;
                }, {});

                // Populate tables for each day
                const columnsToDisplay = [
                    'Student FirstName', 'Student LastName',
                    'Day', 'Start Time', 'End Time',
                    'Parent FirstName', 'Parent LastName',
                    'Phone Number1', 'Phone Number2', 'Email1', 'Email2', 'Date Enrolled'
                ];

                Object.keys(groupedData).forEach((key, index) => {
                    const color = getColor(index);
                    groupedData[key].forEach(item => {
                        const tableType = item['Online'] === 'Y' ? 'Online' : '';
                        const dayTableBody = document.getElementById(`${item.Day}${tableType}Table`).getElementsByTagName('tbody')[0];
                        if (dayTableBody) {
                            const row = document.createElement('tr');
                            row.style.backgroundColor = color;
                            columnsToDisplay.forEach(col => {
                                const cell = document.createElement('td');
                                cell.textContent = item[col] !== undefined ? item[col] : '';
                                row.appendChild(cell);
                            });
                            dayTableBody.appendChild(row);
                        } else {
                            console.error(`Table for ${item.Day}${tableType} not found`);
                        }
                    });
                });

                // Update title
                document.getElementById('mainTitle').textContent = `Active Students and Their Class Schedules`;

                // Tab switching logic
                const tabs = document.querySelectorAll('.tab');
                const tabContents = document.querySelectorAll('.tab-content');

                tabs.forEach(tab => {
                    tab.addEventListener('click', () => {
                        tabs.forEach(t => t.classList.remove('active'));
                        tab.classList.add('active');
                        tabContents.forEach(content => content.classList.remove('active'));
                        document.getElementById(tab.dataset.day).classList.add('active');
                    });
                });

            } catch (error) {
                console.error('Error fetching or parsing file:', error);
            }
        })
        .catch(error => console.error('Error fetching or parsing file:', error));
 });
