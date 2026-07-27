const container = document.getElementById("birthdayContainer");

fetch("birthdays.xlsx")
  .then(response => {
    if (!response.ok) {
      throw new Error("Could not load birthdays.xlsx");
    }
    return response.arrayBuffer();
  })
  .then(data => {

    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const people = XLSX.utils.sheet_to_json(sheet, { raw: true });

    const today = new Date();

    people.forEach(person => {

      let birthday;

      // Handle Excel date values
      if (typeof person.Birthday === "number") {
        const excelDate = XLSX.SSF.parse_date_code(person.Birthday);
        birthday = new Date(
          excelDate.y,
          excelDate.m - 1,
          excelDate.d
        );
      } else {
        birthday = new Date(person.Birthday);
      }

      const isToday =
        birthday.getDate() === today.getDate() &&
        birthday.getMonth() === today.getMonth();

      const card = document.createElement("div");
      card.className = isToday ? "card today" : "card";

      card.innerHTML = `
        <img src="${person.Photo}"
             alt="${person.Name}"
             style="width:150px;height:150px;border-radius:50%;object-fit:cover;"
             onerror="this.src='https://via.placeholder.com/150?text=No+Photo'">

        <h2>${person.Name}</h2>

        <p><strong>Birthday:</strong>
        ${birthday.toLocaleDateString()}</p>

        <p>${person.Message}</p>

        ${isToday ? "<h2>🎉 HAPPY BIRTHDAY! 🎂🎈</h2>" : ""}
      `;

      container.appendChild(card);

    });

  })
  .catch(error => {
    console.error(error);

    container.innerHTML = `
      <h2 style="color:red;text-align:center;">
        Unable to load birthdays.xlsx
      </h2>
    `;
  });
