/*
    JS code:
1. load excel file by file chooser
2. read excel file data
3. create html table from array
*/

var extra_nivele_count = 0;
var original_ds = [];
var compacted_ds = [];
var optimal_ds = [];
var combinatii_unice = [];
var selected_ds = [];
var ds_antemasuratori = [];
var trees = [];
var nivele = [];
var _nivel_oferta = "NIVEL_OFERTA_";
var _nivel_cantitate_articol_oferta = "CANTITATE_ARTICOL_OFERTA";
var _cantitate_antemasuratori = "CANTITATE_ARTICOL_ANTEMASURATORI";

// 1. load excel file by file chooser xlsx.js
function loadDataFromFile(evt) {
  var file = document.getElementById("file_oferta_initiala").files[0];
  var reader = new FileReader();
  reader.onload = function (e) {
    var excel_object = null;
    original_ds = [];
    compacted_ds = [];
    optimal_ds = [];
    combinatii_unice = [];
    var data = e.target.result;
    var workbook = XLSX.read(data, {
      type: "binary",
    });
    workbook.SheetNames.forEach(function (sheetName) {
      var XL_row_object = XLSX.utils.sheet_to_row_object_array(
        workbook.Sheets[sheetName]
      );
      excel_object = JSON.stringify(XL_row_object);
    });
    //console.log("excel_object", excel_object);

    //each key represents a column, copy all objecys in compacted_ds but remove the ones with empty values
    original_ds = JSON.parse(excel_object);
    console.log("original_ds", original_ds);

    compacted_ds = removeEmpty(original_ds);
    //console.log("compacted_ds", compacted_ds);

    const unique_key = "SERIE_ARTICOL_OFERTA";
    optimal_ds = compactByUniqueKey(compacted_ds, unique_key);
    console.log("optimal_ds", optimal_ds);

    pushDataToTable(
      optimal_ds,
      "thead_oferta_initiala",
      "tbody_oferta_initiala"
    );

    var delimiter = "~~~~~~~~~~~~~~~";
    var result = creazaIerarhii(optimal_ds, delimiter);
    nivele = result.nivele;
    var combinatii_unice_as_str = result.combinatii_unice_as_str;
    combinatii_unice = result.combinatii_unice;

    populateSelect(nivele, combinatii_unice_as_str, optimal_ds, delimiter);

    createGraphs(combinatii_unice);
  };

  reader.onerror = function (ex) {
    console.log(ex);
  };

  reader.readAsBinaryString(file);

  var btn_oferta = document.getElementById("btn_oferta");
  btn_oferta.innerHTML = "Salveaza oferta initiala";
  btn_oferta.classList.remove("btn-success");
  btn_oferta.classList.add("btn-danger");
}

function removeEmpty(original_ds) {
  var keys = Object.keys(original_ds[0]);
  var compacted_ds = [];
  original_ds.forEach(function (object) {
    var obj = {};
    keys.forEach(function (key) {
      if (object[key]) {
        obj[key] = object[key];
      }
    });
    compacted_ds.push(obj);
  });

  return compacted_ds;
}

function compactByUniqueKey(compacted_ds, unique_key) {
  //return [...array.reduce((r, o) => r.set(o[key], o), new Map()).values()];

  //rearrange data so all objects with same key unique_key are displayed together
  var optimal_ds = [];
  var distinct = [...new Set(compacted_ds.map((x) => x[unique_key]))];
  distinct.forEach(function (item) {
    compacted_ds.forEach(function (object) {
      if (object[unique_key] == item) {
        optimal_ds.push(object);
      }
    });
  });

  return optimal_ds;
}

function creazaIerarhii(optimal_ds, delimiter = " ") {
  /*
NIVEL_OFERTA_1  NIVEL_OFERTA_2 NIVEL_OFERTA_3 ... NIVEL_OFERTA_N
SUPRATERAN  INSTALATII ELECTRICE	DISTRIBUTIE
SUBSOLURI	INSTALATII ELECTRICE	DISTRIBUTIE   
SUBSOLURI	INSTALATII ELECTRICE	DISTRIBUTIE  

=>

distincy combinations of values for combo NIVEL_OFERTA_1, NIVEL_OFERTA_2, NIVEL_OFERTA_3 ... NIVEL_OFERTA_N
SUPRATERAN  INSTALATII ELECTRICE	DISTRIBUTIE
SUBSOLURI	INSTALATII ELECTRICE	DISTRIBUTIE 
*/

  var nivele = [];
  keys = Object.keys(optimal_ds[0]);
  keys.forEach(function (key) {
    if (key.includes("NIVEL_OFERTA")) {
      nivele.push(key);
    }
  });

  console.log("nivele detectate", nivele);

  //scan optimal_ds and create distinct combinations of values for combo NIVEL_OFERTA_1, NIVEL_OFERTA_2, NIVEL_OFERTA_3 ... NIVEL_OFERTA_N
  var all_combos_for_nivele = [];
  optimal_ds.forEach(function (object) {
    var combo = [];
    nivele.forEach(function (nivel) {
      //if trimmed object[nivel] is not empty add it to combo
      if (object[nivel] && object[nivel].trim()) combo.push(object[nivel]);
    });
    all_combos_for_nivele.push(combo);
  });

  //console.log("all_combos_for_nivele", all_combos_for_nivele);

  //remove duplicates from all_combos_for_nivele
  var combinatii_unice_as_str = [];
  all_combos_for_nivele.forEach(function (combo) {
    const combo_str = combo.join(delimiter);
    if (!combinatii_unice_as_str.includes(combo_str)) {
      combinatii_unice_as_str.push(combo_str);
    }
  });

  //console.log("combinatii_unice_as_str", combinatii_unice_as_str);

  var combinatii_unice = [];
  combinatii_unice_as_str.forEach(function (combo_str) {
    combinatii_unice.push(combo_str.split(delimiter));
  });

  //console.log("combinatii_unice", combinatii_unice);

  return { nivele, combinatii_unice_as_str, combinatii_unice };
}

function populateSelect(
  nivele,
  combinatii_unice_as_str,
  optimal_ds,
  delimiter
) {
  //add combinatii_unice_as_str to select id="ierarchii"
  var select = document.getElementById("ierarhii");
  //add default option
  var option = document.createElement("option");
  option.value = "1";
  option.text = "Toate ierarhiile";
  select.appendChild(option);
  //add combinatii_unice_as_str as options
  combinatii_unice_as_str.forEach(function (combo_str) {
    var option = document.createElement("option");
    option.value = combo_str;
    //option.text = combo_str;
    //replace delimiter with " -> "
    option.text = combo_str.split(delimiter).join(" - ");
    select.appendChild(option);
    select.onchange = function () {
      selected_ds = [];
      if (select.value == "1") {
        pushDataToTable(
          optimal_ds,
          "thead_oferta_initiala",
          "tbody_oferta_initiala"
        );
        return;
      }

      filterOptimalDs(select.value, optimal_ds, delimiter);

      //create table rows
      if (selected_ds.length > 0) {
        pushDataToTable(
          selected_ds,
          "thead_oferta_initiala",
          "tbody_oferta_initiala"
        );
      } else {
        //display a message in table
        var tbody = document.getElementById("tbody_oferta_initiala");
        tbody.innerHTML = "";
        var tr = document.createElement("tr");
        var td = document.createElement("td");
        td.innerHTML = "Nu exista date pentru ierarhia selectata";
        tr.appendChild(td);
        tbody.appendChild(tr);
      }

      drawModalDialog(select.value.split(delimiter), selected_ds);
    };
  });
}

function filterOptimalDs(selected_option, optimal_ds, delimiter) {
  //filter optimal_ds by selected option and display it in table
  optimal_ds.forEach(function (object) {
    var combo = [];
    nivele.forEach(function (nivel) {
      combo.push(object[nivel]);
    });
    if (combo.join(delimiter) == selected_option) {
      selected_ds.push(object);
    } else {
      if (selected_option.includes(combo.join(delimiter))) {
        selected_ds.push(object);
      }
    }
  });

  console.log("selected_ds", selected_ds);
}

function createGraphs(combinatii_unice) {
  //create a graph for each root

  //sort combinatii_unice by first element
  combinatii_unice.sort(function (a, b) {
    return a[0] > b[0] ? 1 : -1;
  });

  console.log("combinatii_unice sorted", combinatii_unice);

  //delete empty elements
  combinatii_unice = combinatii_unice.filter(function (el) {
    return el.length != 0;
  });

  //delete elements with empty strings
  combinatii_unice = combinatii_unice.filter(function (el) {
    return el != "";
  });

  var tree = [];
  var roots = [];
  //find unique roots
  combinatii_unice.forEach(function (combo) {
    if (!roots.includes(combo[0])) {
      roots.push(combo[0]);
    }
  });
  console.log("roots", roots);

  //create trees by roots
  roots.forEach(function (root) {
    tree = [];
    combinatii_unice.forEach(function (combo) {
      if (combo[0] == root) {
        tree.push(combo);
      }
    });
    trees.push(tree);
  });

  console.log("trees", trees);

  //create a new select with tree roots
  var select = document.createElement("select");
  select.id = "select_tree";
  select.classList.add("form-select");
  select.classList.add("form-select-sm");
  //add options
  trees.forEach(function (tree, index) {
    var option = document.createElement("option");
    option.value = index;
    option.text = tree[0][0];
    select.appendChild(option);
  });

  //no selected value
  select.selectedIndex = -1;

  //add select to div id="cytoscape_graphs"
  var div = document.getElementById("select_graphs");
  div.appendChild(select);
  //on change select value create graph
  select.onchange = function () {
    var index = select.value;
    var tree = trees[index];
    var id = document.getElementById("cytoscape_graphs").id;
    createGraph(tree, id);
  };

  function createGraph(tree, id) {
    console.log("selected tree", tree);
    //create a graph for tree
    var cy = cytoscape({
      container: document.getElementById(id),
      elements: elementify(tree),
      style: [
        {
          selector: "node",
          style: {
            "background-color": "#2ecc71",
            label: "data(label)",
          },
        },
        {
          selector: "edge",
          style: {
            "line-color": "#95a5a6",
          },
        },
      ],
    });

    cy.on("cxttapstart", "node", (e) => {
      var node = e.target;
      console.log("node id=" + node.id());
      //find elements[id].label
      var elements = cy.elements();
      var label = elements
        .filter((el) => el.id() == node.id())[0]
        .data("label");
      console.log("label", label);
      console.log("tree", tree);
      // 1. check if it's a leaf
      var isLeaf = false;
      var myBranch;
      for (var i = 0; i < tree.length; i++) {
        var branch = tree[i];
        if (branch[branch.length - 1] == label) {
          isLeaf = true;
          myBranch = branch;
          break;
        }
      }
      console.log(isLeaf, myBranch);

      if (isLeaf && myBranch) {
        var newLeaf = prompt("New leaf\ncam[1..10] or simple string");
        if (newLeaf) {
          //can be "XXX[1..10]" or simple string
          var regex = /\[(.*?)\]/;
          var found = newLeaf.match(regex);
          console.log("found", found);
          if (found) {
            //replace [1..10] with numbers
            var numbers = found[1].split("..");
            var start = parseInt(numbers[0]);
            var end = parseInt(numbers[1]);
            var range = Array.from(
              { length: end - start + 1 },
              (_, i) => i + start
            );
            console.log("range", range);
            range.forEach(function (number) {
              var leaf = newLeaf.replace(found[0], number);
              var cl = [...myBranch];
              cl.push(leaf);
              tree.push(cl);
            });
          } else {
            var cl = [...myBranch];
            cl.push(newLeaf);
            tree.push(cl);
          }
        }

        //redraw graph
        var id = document.getElementById("cytoscape_graphs").id;
        createGraph(tree, id);
      }

      console.log(tree);
    });

    cy.layout({
      name: "breadthfirst",
      fit: true,
      directed: true,
      padding: 10,
      circle: false,
      spacingFactor: 2,
      avoidOverlap: true,
      grid: false,
      // a sorting function to order nodes at equal depth. e.g. function(a, b){ return a.data('weight') - b.data('weight') }
      depthSort: function (a, b) {
        return a.data("rank") - b.data("rank");
      },
      //ready: undefined, // callback on layoutready
      ready: function () {
        console.log("layoutready:" + id);
      },
    }).run();
  }
}

function elementify(combinatii_unice) {
  /*
 [
    [
        "SUPRATERAN",
        "INSTALATII ELECTRICE",
        "DISTRIBUTIE"
    ],
    [
        "SUPRATERAN",
        "INSTALATII ELECTRICE",
        "DISTRIBUTIE",
        "ETAJ 1"
    ],
    [
        "SUPRATERAN",
        "INSTALATII ELECTRICE",
        "DISTRIBUTIE",
        "ETAJ 2"
    ],
    [
        "SUBSOLURI",
        "INSTALATII ELECTRICE",
        "DISTRIBUTIE"
    ]
]

create tree branches: SUPRATERAN -> INSTALATII ELECTRICE -> DISTRIBUTIE -> ETAJ 1
                                                                        -> ETAJ 2
  SUBSOLURI -> INSTALATII ELECTRICE -> DISTRIBUTIE 
  
  every root has it's own branch
 */

  var elements = [],
    nodes = [],
    edges = [],
    rank = 0,
    combinatii_unice_with_id = [];

  combinatii_unice.forEach((combo) => {
    combo_id = [];
    combo.forEach((nivel) => {
      var i = combinatii_unice.indexOf(combo) + 1;
      var j = combo.indexOf(nivel) + 1;
      var id = "n" + i + j;
      //if nivel already exists in combinatii_unice_with_id/combo then get it's id
      //don't look for the last element of combo
      if (j < combo.length) {
        var found = combinatii_unice_with_id
          .flat()
          .find((o) => o.label === nivel);
        if (found) {
          id = found.id;
          //console.log("found", found);
        }
      }
      var label = nivel;
      var o = { id, label };
      combo_id.push(o);
    });
    combinatii_unice_with_id.push(combo_id);
  });

  console.log("combinatii_unice_with_id", combinatii_unice_with_id);

  combinatii_unice_with_id.forEach((combo) => {
    //console.log("combo", combo);
    rank = 0;
    var parent = null;
    combo.forEach((nivel) => {
      rank++;
      var node = { data: { id: nivel.id, label: nivel.label, rank } };
      if (parent) {
        var edge = { data: { source: parent.data.id, target: node.data.id } };
        edges.push(edge);
      }
      nodes.push(node);
      parent = node;
    });
  });

  //remove duplicates from nodes and edges
  nodes = nodes.filter(
    (v, i, a) => a.findIndex((t) => t.data.id === v.data.id) === i
  );
  edges = edges.filter(
    (v, i, a) =>
      a.findIndex(
        (t) =>
          t.data.source === v.data.source && t.data.target === v.data.target
      ) === i
  );

  elements = nodes.concat(edges);

  //sort elements by rank
  elements.sort(function (a, b) {
    return a.data.rank - b.data.rank;
  });

  console.log("elements", elements);

  return elements;
}

function pushDataToTable(data, thead_name, tbody_name) {
  //create html table from array
  //data > table id="table_oferta_initiala"
  //thead id="thead_oferta_initiala" < first object's keys
  //tbody id="tbody_oferta_initiala" < all objects as rows
  var thead = document.getElementById(thead_name);
  thead.innerHTML = "";
  var tbody = document.getElementById(tbody_name);
  tbody.innerHTML = "";
  //add delete icon to thead
  var th = document.createElement("th");
  th.innerHTML = "";
  thead.appendChild(th);
  var keys = Object.keys(data[0]);
  keys.forEach(function (key) {
    var th = document.createElement("th");
    th.style.writingMode = "vertical-rl";
    th.style.rotate = "180deg";
    /* //sticky top
    th.style.position = "sticky";
    th.style.top = "0";
    th.style.backgroundColor = "white"; */
    th.innerHTML = key;
    //add class header
    th.classList.add("header");
    //add scope col
    th.setAttribute("scope", "col");
    thead.appendChild(th);
  });

  //create table rows
  data.forEach(function (object) {
    var tr = document.createElement("tr");
    //add td with delete icon
    var td = document.createElement("td");
    var icon = document.createElement("i");
    icon.classList.add("bi");
    icon.classList.add("bi-trash");
    icon.classList.add("text-danger");
    icon.style.cursor = "pointer";
    icon.onclick = function () {
      //delete row
    };
    td.appendChild(icon);
    tr.appendChild(td);
    keys.forEach(function (key) {
      var td = document.createElement("td");
      var val = 0;
      //if object[key] is a number format it to 2 decimals
      if (object[key]) {
        val = object[key];
        if (!isNaN(val)) {
          val = parseFloat(val).toFixed(2);
        } else {
          val = object[key];
        }
      }

      td.innerHTML = val || "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}

function drawModalDialog(selected_combo, selected_ds) {
  extra_nivele_count = 0;
  console.log("selected_combo", selected_combo);
  //add to modal id="modal-body"; 1. split combo by delimiter  2.add each part as a text input
  var modal_body = document.getElementById("modal-body");
  var modal_header = document.getElementById("modal-header");
  modal_body.innerHTML = "";
  modal_header.innerHTML = "";
  //add div container
  var container = document.createElement("div");
  container.classList.add("container");
  modal_header.appendChild(container);
  //create div for all selected_combo
  var row = document.createElement("div");
  row.classList.add("row");
  row.id = "nivele_ofertare";
  container.appendChild(row);
  selected_combo.forEach(function (part) {
    var niv = document.createElement("h6");
    niv.id = _nivel_oferta + (selected_combo.indexOf(part) + 1);
    niv.classList.add("col-sm");
    niv.classList.add("text-primary");
    niv.innerHTML = part;
    row.appendChild(niv);
  });

  //add all selected_ds['DENUMIRE_ARTICOL_OFERTA'] to modal id="modal-body" as a newly created select
  //create a div for select, label and button
  var row1 = document.createElement("div");
  row1.classList.add("row");
  row1.id = "articole_antemasuratori";
  container.appendChild(row1);
  var div2 = document.createElement("div");
  div2.classList.add("col-sm-8");
  var select_articole = document.createElement("select");
  select_articole.id = "articole";
  select_articole.name = "articole";
  select_articole.classList.add("form-select");
  select_articole.classList.add("form-select-sm");
  //multiselect
  select_articole.multiple = true;
  //8 rows
  select_articole.size = 8;
  selected_ds.forEach(function (object) {
    var option = document.createElement("option");
    option.value = object["WBS"];
    option.text =
      object["DENUMIRE_ARTICOL_OFERTA"] +
      " - " +
      object["CANTITATE_ARTICOL_OFERTA"] +
      " [" +
      object["UM_ARTICOL_OFERTA"] +
      "]";
    select_articole.appendChild(option);
  });
  div2.appendChild(select_articole);
  row1.appendChild(div2);

  //add button Adauga in antemasuratori
  var div_btn = document.createElement("div");
  div_btn.classList.add("col-sm");
  var button = document.createElement("button");
  button.type = "button";
  button.classList.add("btn");
  button.classList.add("btn-primary");
  button.classList.add("btn-sm");
  var arrow_down =
    '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" class="bi bi-arrow-down" viewBox="0 0 16 16">' +
    '<path fill-rule="evenodd" d="M8 1a.5.5 0 0 1 .5.5v11.793l3.146-3.147a.5.5 0 0 1 .708.708l-4 4a.5.5 0 0 1-.708 0l-4-4a.5.5 0 0 1 .708-.708L7.5 13.293V1.5A.5.5 0 0 1 8 1"></path>' +
    "</svg>";
  button.innerHTML = "Antemasuratori" + arrow_down;
  div_btn.appendChild(button);
  row1.appendChild(div_btn);
  button.onclick = function () {
    //adauga in ds_antemasuratori
    var selected_articole = [];
    for (var i = 0; i < select_articole.options.length; i++) {
      if (select_articole.options[i].selected) {
        selected_articole.push(select_articole.options[i].value);
      }
    }
    console.log("selected_articole", selected_articole);
    //find selected WBS  in selected_ds['WBS'] => a new ds with only selected WBS
    selected_ds.forEach(function (object) {
      if (selected_articole.includes(object["WBS"])) {
        ds_antemasuratori.push(object);
      }
    });
    console.log("ds_antemasuratori", ds_antemasuratori);

    var new_levels = [];
    for (var i = 0; i < selected_combo.length; i++) {
      new_levels.push(_nivel_oferta + (i + 1));
    }

    console.log("new_levels", new_levels);

    //enumarate keys din ds_antemasuratori
    var cut1,
      cut2,
      keys = Object.keys(ds_antemasuratori[0]);
    console.log("keys", keys);
    for (var i = 0; i < keys.length; i++) {
      if (keys[i].includes(_nivel_oferta)) {
        cut1 = i;
      }
      if (keys[i].includes(_nivel_cantitate_articol_oferta)) {
        if (!document.getElementById("tbody_antemasuratori")) {
          cut2 = i;
        } else {
          cut2 = i - 1;
        }
      }
    }

    console.log("cut1", cut1);
    console.log("cut1", cut2);

    var rearranged_ds_antemasuratori = [];
    //=> ds_antemasuratori keys: 0 ... cut1 ramains the same + add unique new_levels: selected_combo + add remaining cut1+1 ... end
    ds_antemasuratori.forEach(function (object) {
      var new_object = {};
      for (var i = 0; i < cut1 + 1; i++) {
        new_object[keys[i]] = object[keys[i]];
      }
      for (var i = 0; i < new_levels.length; i++) {
        new_object[new_levels[i]] = selected_combo[i].toString();
      }
      for (var i = cut1 + 1; i <= cut2; i++) {
        new_object[keys[i]] = object[keys[i]];
      }
      new_object[
        _cantitate_antemasuratori
      ] = `<input type="text" class="form-control form-control-sm bg-warning cantitate_articol_antemasuratori" value="0" />`;
      for (var i = cut2 + 1; i < keys.length; i++) {
        new_object[keys[i]] = object[keys[i]];
      }
      rearranged_ds_antemasuratori.push(new_object);
    });

    ds_antemasuratori = [...rearranged_ds_antemasuratori];

    console.log("ds_antemasuratori", ds_antemasuratori);

    //create table
    var table = document.createElement("table");
    table.classList.add("table");
    table.classList.add("table-sm");
    table.classList.add("table-bordered");
    table.classList.add("table-hover");
    table.classList.add("table-striped");
    table.classList.add("table-responsive");
    table.id = "table_antemasuratori";
    //get or create thead and tbody antemasuratori
    var thead_antemasuratori =
      document.getElementById("thead_antemasuratori") ||
      document.createElement("thead");
    thead_antemasuratori.id = "thead_antemasuratori";
    var tbody_antemasuratori =
      document.getElementById("tbody_antemasuratori") ||
      document.createElement("tbody");
    tbody_antemasuratori.id = "tbody_antemasuratori";
    table.appendChild(thead_antemasuratori);
    table.appendChild(tbody_antemasuratori);
    modal_body.appendChild(table);
    pushDataToTable(
      ds_antemasuratori,
      "thead_antemasuratori",
      "tbody_antemasuratori"
    );
  };
}

function saveChanges() {
  document.getElementById("modal-body").innerHTML = "";
  //close modal
  var modal = new bootstrap.Modal(
    document.getElementById("AntemasuratoriModal")
  );
  modal.hide();
}

function saveOferta() {
  var btn_oferta = document.getElementById("btn_oferta");
  btn_oferta.innerHTML = "Oferta salvata";
  btn_oferta.classList.remove("btn-danger");
  btn_oferta.classList.add("btn-success");
  alert("to be implemented");
}

function populateSelectIerarhiiFromTrees() {
  var select = document.getElementById("ierarhii");
  select.innerHTML = "";
  //add default option
  var option = document.createElement("option");
  option.value = "1";
  option.text = "Toate ierarhiile";
  select.appendChild(option);
  trees.forEach(function (tree, index) {
    tree.forEach(function (branch) {
      var option = document.createElement("option");
      option.value = branch.join("~~~~~~~~~~~~~~~");
      option.text = branch.join(" - ");
      select.appendChild(option);
    });
  });
}

//add to .cantitate_articol_antemasuratori on arrow up/down => move to next/previous input
document.addEventListener("keydown", function (e) {
  if (e.key === "ArrowUp" || e.key === "ArrowDown") {
    var inputs = document.querySelectorAll(".cantitate_articol_antemasuratori");
    var index = Array.from(inputs).indexOf(document.activeElement);
    if (index > -1) {
      if (e.key === "ArrowUp") {
        if (index > 0) {
          inputs[index - 1].focus();
        }
      } else if (e.key === "ArrowDown") {
        if (index < inputs.length - 1) {
          inputs[index + 1].focus();
        }
      }
    }
  }
});

//add to .cantitate_articol_antemasuratori on enter => move to next input
document.addEventListener("keydown", function (e) {
  if (e.key === "Enter") {
    var inputs = document.querySelectorAll(".cantitate_articol_antemasuratori");
    var index = Array.from(inputs).indexOf(document.activeElement);
    if (index > -1) {
      if (index < inputs.length - 1) {
        inputs[index + 1].focus();
      }
    }
  }
});

//add to .cantitate_articol_antemasuratori on focus => select all text
document.addEventListener("focus", function (e) {
  if (e.target.classList.contains("cantitate_articol_antemasuratori")) {
    e.target.select();
  }
});

//add to .cantitate_articol_antemasuratori on input => ds_antemasuratori[index][CANTITATE_ARTICOL_ANTEMASURATORI] = input.value
document.addEventListener("input", function (e) {
  if (e.target.classList.contains("cantitate_articol_antemasuratori")) {
    var index = Array.from(
      document.querySelectorAll(".cantitate_articol_antemasuratori")
    ).indexOf(e.target);
    console.log("index", index);
    ds_antemasuratori[index][_cantitate_antemasuratori] = parseFloat(
      e.target.value
    );
    console.log("ds_antemasuratori", ds_antemasuratori);
  }
});
