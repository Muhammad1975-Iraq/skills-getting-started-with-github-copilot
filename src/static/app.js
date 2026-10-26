document.addEventListener("DOMContentLoaded", () => {
  const activitiesList = document.getElementById("activity-grid");
  const summaryCards = document.getElementById("summary-cards");
  const activitySelect = document.getElementById("activity");
  const signupForm = document.getElementById("signup-form");
  const messageDiv = document.getElementById("message");
  const recentSignupsList = document.getElementById("recent-signups");
  const downloadExcelButton = document.getElementById("download-excel");

  let dashboardData = { summary: {}, activities: {} };

  function buildMetricCard(label, value, meta) {
    return `
      <div class="metric-card">
        <span class="metric-label">${label}</span>
        <span class="metric-value">${value}</span>
        <span class="metric-meta">${meta}</span>
      </div>
    `;
  }

  function renderSummary(summary = {}) {
    const cards = [
      buildMetricCard("Activities", summary.total_activities ?? 0, "Active programs"),
      buildMetricCard("Participants", summary.total_participants ?? 0, "Students enrolled"),
      buildMetricCard("Open Slots", summary.open_slots ?? 0, `${summary.utilization ?? 0}% utilization`),
      buildMetricCard("Capacity", summary.capacity ?? 0, "Total seats")
    ];

    summaryCards.innerHTML = cards.join("");
  }

  function renderRecentSignups(activities = {}) {
    const recent = Object.entries(activities)
      .flatMap(([name, details]) =>
        (details.participants || []).map((email) => ({ name, email }))
      )
      .slice(-5)
      .reverse();

    if (!recent.length) {
      recentSignupsList.innerHTML = "<li><span>No recent signups</span></li>";
      return;
    }

    recentSignupsList.innerHTML = recent
      .map(
        ({ name, email }) => `
          <li>
            <strong>${name}</strong>
            <span>${email}</span>
          </li>
        `
      )
      .join("");
  }

  function renderActivities(activities = {}) {
    const activityCards = Object.entries(activities)
      .map(([name, details]) => {
        const spotsLeft = details.max_participants - (details.participants || []).length;
        const participants = (details.participants || []).length
          ? (details.participants || [])
              .map((email) => `<li>${email}</li>`)
              .join("")
          : "<li>No participants yet</li>";

        return `
          <article class="activity-card">
            <h3>${name}</h3>
            <p>${details.description}</p>
            <p><strong>Schedule:</strong> ${details.schedule}</p>
            <div class="activity-meta">
              <span>${spotsLeft} seats left</span>
              <span>${details.max_participants} max</span>
            </div>
            <ul class="participants-list">${participants}</ul>
          </article>
        `;
      })
      .join("");

    activitiesList.innerHTML = activityCards;
  }

  function populateActivitySelect(activities = {}) {
    activitySelect.innerHTML = '<option value="">-- Select an activity --</option>';

    Object.keys(activities).forEach((name) => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      activitySelect.appendChild(option);
    });
  }

  async function fetchActivities() {
    try {
      const response = await fetch("/dashboard");
      const data = await response.json();
      dashboardData = data;
      const activities = data.activities || {};
      const summary = data.summary || {};

      renderSummary(summary);
      renderActivities(activities);
      populateActivitySelect(activities);
      renderRecentSignups(activities);
    } catch (error) {
      activitiesList.innerHTML = "<p>Failed to load activities. Please try again later.</p>";
      console.error("Error fetching activities:", error);
    }
  }

  function showMessage(text, type = "success") {
    messageDiv.textContent = text;
    messageDiv.className = `message ${type}`;
    messageDiv.classList.remove("hidden");

    setTimeout(() => {
      messageDiv.classList.add("hidden");
    }, 5000);
  }

  function exportDashboardToExcel() {
    const summary = dashboardData.summary || {};
    const activities = dashboardData.activities || {};

    const summaryRows = [
      ["Metric", "Value"],
      ["Total Activities", summary.total_activities ?? 0],
      ["Total Participants", summary.total_participants ?? 0],
      ["Open Slots", summary.open_slots ?? 0],
      ["Capacity", summary.capacity ?? 0],
      ["Utilization %", summary.utilization ?? 0]
    ];

    const activityRows = [
      ["Activity", "Description", "Schedule", "Max Participants", "Participants Count", "Participants"],
      ...Object.entries(activities).map(([name, details]) => [
        name,
        details.description || "",
        details.schedule || "",
        details.max_participants || 0,
        (details.participants || []).length,
        (details.participants || []).join("; ")
      ])
    ];

    const workbook = XLSX.utils.book_new();

    const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
    XLSX.utils.book_append_sheet(workbook, summarySheet, "Summary");

    const activitySheet = XLSX.utils.aoa_to_sheet(activityRows);
    XLSX.utils.book_append_sheet(workbook, activitySheet, "Activities");

    XLSX.writeFile(workbook, "rixios-dashboard.xlsx");
  }

  signupForm.addEventListener("submit", async (event) => {
    event.preventDefault();

    const email = document.getElementById("email").value;
    const activity = document.getElementById("activity").value;

    try {
      const response = await fetch(
        `/activities/${encodeURIComponent(activity)}/signup?email=${encodeURIComponent(email)}`,
        {
          method: "POST",
        }
      );

      const result = await response.json();

      if (response.ok) {
        showMessage(result.message, "success");
        signupForm.reset();
        await fetchActivities();
      } else {
        showMessage(result.detail || "An error occurred", "error");
      }
    } catch (error) {
      showMessage("Failed to sign up. Please try again.", "error");
      console.error("Error signing up:", error);
    }
  });

  downloadExcelButton.addEventListener("click", exportDashboardToExcel);

  fetchActivities();
});
