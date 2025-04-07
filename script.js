const sheetConfigs = {
  "All Tests": { filters: ["Version:", "Type:", "Study", "Source"] },
  "ABC": { filters: ["Version:", "Questionaire:", "Dictionary:", "Dictionary of Values:", "Other:"] },
  "EHS": { filters: ["Version:", "Questionaire:", "Dictionary:", "Dictionary of Values:", "Other:"] },
  "HSIS": { filters: ["Version:", "Questionaire:", "Dictionary:", "Dictionary of Values:", "Other:"] },
  "Perry": { filters: ["Version:", "Questionaire:", "Dictionary:", "Dictionary of Values:", "Other:", "More:"] },
  "HS FACES": { filters: ["Version:", "Questionaire:", "Scale Construction:", "Dictionary:", "Other:"] },
  "NFP-M": { filters: ["Version:", "Questionaire:", "Dictionary:", "Dictionary of Values:"] },
};

fetch("Wiki_AH.xlsx")
  .then(res => res.arrayBuffer())
  .then(ab => {
    const workbook = XLSX.read(ab, { type: "array" });

    Object.entries(sheetConfigs).forEach(([sheetName, config]) => {
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      if (!data.length) return;

      // Create tab content section
      const container = document.createElement("div");
      container.className = `tab-pane fade${sheetName === "All Tests" ? " show active" : ""}`;
      container.id = sheetName.replace(/[^a-zA-Z0-9]/g, '_');

      // Filters section
      const filterDiv = document.createElement("div");
      filterDiv.className = "filters-section d-flex flex-wrap gap-3 mb-3";
      const filterInputs = {};

      config.filters.forEach(filter => {
        if (filter === "Study") {
          const group = document.createElement("div");
          group.innerHTML = `<label class='form-label d-block'>Study</label>
            <div class='d-flex flex-wrap gap-2'>
              ${["ABC","Perry","HSIS","NFP","IHDP","EHS"].map(s => `
                <div class='form-check'>
                  <input type='checkbox' class='form-check-input' id='${sheetName}_study_${s}' value='${s}'>
                  <label class='form-check-label' for='${sheetName}_study_${s}'>${s}</label>
                </div>
              `).join("")}
            </div>`;
          filterInputs["Study"] = () => {
            return ["ABC","Perry","HSIS","NFP","IHDP","EHS"].filter(s =>
              document.getElementById(`${sheetName}_study_${s}`)?.checked
            );
          };
          filterDiv.appendChild(group);
        } else {
          const select = document.createElement("select");
          select.className = "form-select";
          select.innerHTML = `<option value=''>-- Show All --</option>` +
            [...new Set(data.map(row => row[filter]).filter(Boolean))]
              .map(val => `<option value="${val}">${val}</option>`).join("");
          const group = document.createElement("div");
          group.innerHTML = `<label class='form-label'>${filter.replace(":", "")}</label>`;
          group.appendChild(select);
          filterInputs[filter] = () => select.value;
          filterDiv.appendChild(group);
        }
      });

      container.appendChild(filterDiv);

      // Table output
      const tableWrapper = document.createElement("div");
      tableWrapper.className = "table-wrapper";
      tableWrapper.id = `${sheetName.replace(/[^a-zA-Z0-9]/g, '_')}_table_wrapper`;

      const table = document.createElement("table");
      table.innerHTML = "<thead></thead><tbody></tbody>";
      table.id = `${sheetName.replace(/[^a-zA-Z0-9]/g, '_')}_table`;
      tableWrapper.appendChild(table);
      container.appendChild(tableWrapper);

      document.getElementById("sheetContent").appendChild(container);

      // Scrollbar
      const scrollBar = document.createElement("div");
      scrollBar.className = "custom-scrollbar-container";
      scrollBar.id = `${sheetName.replace(/[^a-zA-Z0-9]/g, '_')}_scrollbar`;
      scrollBar.innerHTML = `<div class='custom-scrollbar-track'></div>`;
      document.getElementById("scrollbarsContainer").appendChild(scrollBar);

      // Render table function
      const renderTable = (rows) => {
        const thead = table.querySelector("thead");
        const tbody = table.querySelector("tbody");
        thead.innerHTML = tbody.innerHTML = "";
        if (!rows.length) {
          tbody.innerHTML = "<tr><td colspan='5'>No data to show.</td></tr>";
          return;
        }
        const headers = Object.keys(rows[0]);
        thead.innerHTML = `<tr>${headers.map(h => `<th>${h}</th>`).join("")}</tr>`;
        tbody.innerHTML = rows.map(row =>
          `<tr>${headers.map(h => `<td>${row[h] || ""}</td>`).join("")}</tr>`
        ).join("");

        const track = scrollBar.querySelector(".custom-scrollbar-track");
        track.style.width = `${tableWrapper.scrollWidth}px`;
        scrollBar.style.display = "block";
        scrollBar.onscroll = () => tableWrapper.scrollLeft = scrollBar.scrollLeft;
        tableWrapper.onscroll = () => scrollBar.scrollLeft = tableWrapper.scrollLeft;
      };

      // Filtering
      const applyFilters = () => {
        const versioned = data.filter(row => {
          return config.filters.every(f => {
            if (f === "Study") {
              const studies = filterInputs[f]();
              return studies.length === 0 || studies.some(s => row[s]?.toString().trim() !== "");
            } else {
              const selected = filterInputs[f]();
              return !selected || row[f] === selected;
            }
          });
        });
        renderTable(versioned);
      };

      Object.keys(filterInputs).forEach(f => {
        if (f === "Study") {
          ["ABC", "Perry", "HSIS", "NFP", "IHDP", "EHS"].forEach(s => {
            const el = document.getElementById(`${sheetName}_study_${s}`);
            if (el) el.addEventListener("change", applyFilters);
          });
        } else {
          const el = filterDiv.querySelector("select");
          if (el) el.addEventListener("change", applyFilters);
        }
      });

      renderTable(data);
    });
  });
