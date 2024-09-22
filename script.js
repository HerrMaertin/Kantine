let persons = [];
const holidays = [
    "01.01.2024", // Neujahr
    "30.03.2024", // Karfreitag
    "01.05.2024", // Tag der Arbeit
    "09.05.2024", // Christi Himmelfahrt
    "20.05.2024", // Pfingstmontag
    "03.10.2024", // Tag der Deutschen Einheit
    "24.12.2024", // 1. Weihnachtsfeiertag
    "25.12.2024", // 1. Weihnachtsfeiertag
    "26.12.2024"  // 2. Weihnachtsfeiertag
];

// Funktion zur Bestimmung des Wochentages
function getWeekday(date) {
    const days = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];
    return days[date.getDay()];
}

function addPerson() {
    const name = document.getElementById("personName").value;
    if (name.trim() === "") {
        alert("Bitte geben Sie einen Namen ein.");
        return;
    }
    
    persons.push({ name: name, absentDays: [], serviceCount: 0 });
    document.getElementById("personName").value = ""; 
    renderPersons();
}

function renderPersons() {
    const tableBody = document.getElementById("personTableBody");
    tableBody.innerHTML = ""; 

    persons.forEach((person, index) => {
        const row = document.createElement("tr");
        
        const nameCell = document.createElement("td");
        nameCell.textContent = person.name;
        row.appendChild(nameCell);

        const absentDaysCell = document.createElement("td");
        const absentDaysInput = document.createElement("input");
        absentDaysInput.type = "text";
        absentDaysInput.placeholder = "z.B. 21.09, 22.09";
        absentDaysInput.value = person.absentDays.map(d => formatDateToDDMMYYYY(new Date(d))).join(", ");
        absentDaysInput.onchange = (e) => {
            const absentDaysList = e.target.value.split(",").map(d => d.trim());
            const parsedAbsentDays = absentDaysList.map(parseAbsentDay).filter(date => date !== null);
            person.absentDays = parsedAbsentDays;
            if (parsedAbsentDays.length !== absentDaysList.length) {
                alert("Einige der eingegebenen Fehltage waren im falschen Format (tt.mm).");
            }
        };
        absentDaysCell.appendChild(absentDaysInput);
        row.appendChild(absentDaysCell);

        const deleteCell = document.createElement("td");
        const deleteButton = document.createElement("button");
        deleteButton.textContent = "Löschen";
        deleteButton.onclick = () => deletePerson(index);
        deleteCell.appendChild(deleteButton);
        row.appendChild(deleteCell);

        tableBody.appendChild(row);
    });
}

function deletePerson(index) {
    persons.splice(index, 1);
    renderPersons();
}

function calcWorkDays() {
    const dutyPlan = [];
    const dayCounts = {};

    persons.forEach(person => dayCounts[person.name] = 0);

    let currentDate = new Date();

    while (currentDate.getFullYear() === new Date().getFullYear()) { 

        const day = currentDate.getDay();
        const currentDateStr = formatDateToDDMMYYYY(currentDate);
        const weekday = getWeekday(currentDate);  // Wochentag berechnen

        const isHoliday = holidays.includes(currentDateStr);
        
        if (day !== 0 && day !== 6 && !isHoliday) {  // 0 = Sunday, 6 = Saturday
            const dateStr = formatDateToDDMMYYYY(currentDate);

            // check available Persons
            const availablePersons = persons.filter(person => !person.absentDays.includes(dateStr));
            
            if (availablePersons.length > 0) {
               // Find the person with the fewest canteen services
                availablePersons.sort((a, b) => a.serviceCount - b.serviceCount);
                const dutyPerson = availablePersons[0]; 
                
                dutyPlan.push(`${dateStr} (${weekday}): ${dutyPerson.name} hat Kantinendienst`);
                
                dutyPerson.serviceCount++;
                dayCounts[dutyPerson.name]++;
            } else {
                dutyPlan.push(`${dateStr} (${weekday}): Niemand ist verfügbar für den Kantinendienst.`);
            }
        } else {
            dutyPlan.push(`${formatDateToDDMMYYYY(currentDate)} (${weekday}): Kein Dienst nötig (Wochenende oder Feiertag)`);
        }
        
        //check next day
        currentDate.setDate(currentDate.getDate() + 1);
    }

    document.getElementById("workplanResult").innerHTML = dutyPlan.join("<br>");

    let frequencyText = "<h3>Häufigkeit des Dienstes:</h3>";
    for (const person in dayCounts) {
        frequencyText += `${person}: ${dayCounts[person]} mal Kantinendienst<br>`;
    }
    document.getElementById("frequencyResult").innerHTML = frequencyText;
}

function exportToExcel() {
    const dutyPlanText = document.getElementById("workplanResult").innerText.split("\n");

    const wb = XLSX.utils.book_new();
    const ws_data = [["Datum", "Wochentag", "Diensthabende Person"]];

    let previousMonth = null;

    dutyPlanText.forEach(row => {
        const [dateWithWeekday, person] = row.split(": ");
        const [date, weekday] = dateWithWeekday.split(" (");

        if (date && person && weekday) {
            const [day, month, year] = date.split(".");
            const dateObj = new Date(`${year}-${month}-${day}`);
            const cleanedWeekday = weekday.replace(")", "");  // Bereinigter Wochentag

            if (previousMonth !== null && previousMonth !== month) {
                ws_data.push([]);
            }
            previousMonth = month;

            ws_data.push([date, cleanedWeekday, person]);
        }
    });

    ws_data.push([]);
    ws_data.push(["Häufigkeit des Dienstes"]);
    const frequencyText = document.getElementById("frequencyResult").innerHTML.split("<br>");
    frequencyText.forEach(row => {
        if (row.trim()) {
            ws_data.push([row]);
        }
    });

    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    XLSX.utils.book_append_sheet(wb, ws, "Kantinendienst");
    XLSX.writeFile(wb, "Kantinendienst_Plan.xlsx");
}

function formatDateToDDMMYYYY(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}.${month}.${year}`;
}

function parseAbsentDay(absentDay) {
    const [day, month] = absentDay.split(".");
    if (!day || !month || isNaN(day) || isNaN(month)) {
        return null;
    }
    const year = new Date().getFullYear();
    const date = new Date(year, month - 1, day);

    if (date.getDate() == day && date.getMonth() + 1 == month) {
        return formatDateToDDMMYYYY(date);
    } else {
        return null;
    }
}
