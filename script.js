function myFunction() {
  document.getElementById("myTopnav").classList.toggle("responsive");

const container = document.getElementById("birthdayContainer");

fetch("family.xlsx")
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
  const currentMonth = today.getMonth();
  const currentDay = today.getDate();

  const todayBirthdays = [];
  const pastBirthdays = [];
  const upcomingBirthdays = [];

  people.forEach(person => {

    let birthday;

    if (typeof person.Birthday === "number") {
      const excelDate = XLSX.SSF.parse_date_code(person.Birthday);
      birthday = new Date(excelDate.y, excelDate.m - 1, excelDate.d);
    } else {
      birthday = new Date(person.Birthday);
    }

    // Only show birthdays in the current month
    if (birthday.getMonth() !== currentMonth) return;

    const day = birthday.getDate();

    if (day === currentDay) {
      todayBirthdays.push({ person, birthday });
    } else if (day < currentDay) {
      pastBirthdays.push({ person, birthday });
    } else {
      upcomingBirthdays.push({ person, birthday });
    }

  });

  // Sort by day
  pastBirthdays.sort((a, b) => b.birthday.getDate() - a.birthday.getDate());
  upcomingBirthdays.sort((a, b) => a.birthday.getDate() - b.birthday.getDate());

  function createSection(title, list) {

    if (list.length === 0) return;

    const heading = document.createElement("h1");
    heading.className = "section-title";
    heading.textContent = title;

    container.appendChild(heading);

    const section = document.createElement("div");
    section.className = "birthday-section";

    list.forEach(({ person, birthday }) => {

      const isToday = birthday.getDate() === currentDay;

      const card = document.createElement("div");
      card.className = isToday ? "card today" : "card";

      card.innerHTML = `
        <div class="photo-wrapper">
          <img
            src="${person.Photo}"
            alt="${person.Name}"
            class="profile-photo"
            onerror="this.src='https://via.placeholder.com/180?text=No+Photo'">
        </div>

        <div class="card-body">

          <h2>${person.Name}</h2>

          <p class="birthday">
            🎂 ${birthday.toLocaleDateString(undefined,{
              month:'long',
              day:'numeric'
            })}
          </p>

          <p class="message">${person.Message || ""}</p>

          ${isToday ? `
            <div class="birthday-banner">
              🎉 HAPPY BIRTHDAY! 🎂
            </div>
          ` : ""}

        </div>
      `;

      section.appendChild(card);

    });

    container.appendChild(section);

  }

  createSection("🎉 Today's Birthdays", todayBirthdays);
  createSection("📅 Earlier This Month", pastBirthdays);
  createSection("🎂 Coming Up This Month", upcomingBirthdays);

})
  .catch(error => {
    console.error(error);

    container.innerHTML = `
      <h2 style="color:red;text-align:center;">
        Unable to load birthdays.xlsx
      </h2>
    `;
  });
