Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    retrieveTemplatesFromTemplatesFolder();
  }
});

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}

async function clearDocument() {
  await tryCatch(async () => {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.clear();
      await context.sync();
    });
  });
}

async function retrieveTemplatesFromTemplatesFolder() {
  const filePath = "/templates/templates.json";
  try {
    const response = await fetch(filePath);

    if (!response.ok) {
      console.error("err:", response.status, response.statusText);
      throw new Error("err " + response.status);
    }

    const templates = await response.json();
    const groupedTemplates = groupTemplatesByCategory(templates);
    createCategorySections(groupedTemplates);
  } catch (error) {
    console.error("err: ", error);
  }
}

function groupTemplatesByCategory(templates) {
  return templates.reduce((groups, template) => {
    if (!groups[template.category]) {
      groups[template.category] = [];
    }
    groups[template.category].push(template);
    return groups;
  }, {});
}


function createCategorySections(groupedTemplates) {
  const buttonContainer = document.getElementById("template-buttons");
  buttonContainer.innerHTML = ''; 

  for (const category in groupedTemplates) {
    const categorySection = document.createElement("details");
    const categorySummary = document.createElement("summary");
    categorySummary.innerText = category;
    categorySection.appendChild(categorySummary);

    const templates = groupedTemplates[category];
    templates.forEach(template => {
      const button = document.createElement("button");
      button.innerText = template.title;
      button.classList.add("template-button");
      button.onclick = () => fetchTemplateAndImport(template.template);
      button.style.margin = "10px 0"; 

      const description = document.createElement("p");
      description.innerText = template.description;
      description.style.margin = "5px 0";
      description.style.fontSize = "0.9em";
      description.style.color = "#555";

      const templateWrapper = document.createElement("div");
      templateWrapper.appendChild(button);
      templateWrapper.appendChild(description);

      categorySection.appendChild(templateWrapper);
    });

    buttonContainer.appendChild(categorySection);
  }
}


async function fetchTemplateAndImport(templateName) {
  const templatePath = `/templates/${templateName}`;  
  try {
    const response = await fetch(templatePath);

    if (!response.ok) {
      console.error(`Err: ${templateName}`, response.status, response.statusText);
      throw new Error("err " + response.status);
    }

    const fileBlob = await response.blob();
    const base64Template = await convertBlobToBase64(fileBlob);
    importTemplate(base64Template);
  } catch (error) {
    console.error("err:", error);
  }
}

function convertBlobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      const base64Data = reader.result.split(',')[1]; 
      resolve(base64Data); 
    };
    reader.onerror = (error) => {
      reject(error); 
    };
    reader.readAsDataURL(blob);
  });
}

async function importTemplate(base64Template) {
  await Word.run(async (context) => {
    context.document.insertFileFromBase64(base64Template, "Replace", {
      importTheme: true,
      importStyles: true,
      importDifferentOddEvenPages: true,
      importPageColor: true,
      importDifferentOddEvenPages: true
    });
    await context.sync();
  });
}


document.getElementById("search-bar").addEventListener("input", function (event) {
  const query = event.target.value.toLowerCase();
  filterTemplates(query);
});

document.getElementById("clear-button").addEventListener("click", function () {
  clearDocument();
});

function filterTemplates(query) {
  const categories = document.querySelectorAll("#template-buttons > details");

  categories.forEach(category => {
    const templateWrappers = category.querySelectorAll("div");
    let categoryHasMatch = false;

    templateWrappers.forEach(wrapper => {
      const button = wrapper.querySelector("button");
      const description = wrapper.querySelector("p");

      const title = button.innerText.toLowerCase();
      const matches = title.includes(query);

      wrapper.style.display = matches ? "block" : "none";

      if (matches) {
        categoryHasMatch = true;
      }
    });

    category.style.display = categoryHasMatch ? "block" : "none";
  });
}
