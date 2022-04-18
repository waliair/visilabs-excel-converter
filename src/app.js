const Toast = Swal.mixin({
  toast: true,
  position: 'top-right',
  iconColor: 'white',
  customClass: {
    popup: 'colored-toast'
  },
  showConfirmButton: false,
  timer: 4000,
  timerProgressBar: true
})

document.getElementById('upload').addEventListener('change', handleFileSelect, false);
    $(document).on("click", ".dropdown-item", handleJSONObject);
    document.getElementById('xl__json').addEventListener("click",getJsonData,false);
    document.getElementById('copyofresult').addEventListener("click",copyAll,false);
    $(document).on("click", "#loadedSheetNames .nav-link", changeSheet);


    var stateObject = {};

    var ExcelToJSON = function() {

      this.parseExcel = function(file) {
        var reader = new FileReader();
        clearForNewExcel();
        reader.onload = function(e) {
          var data = e.target.result;
          var workbook = XLSX.read(data, {
            type: 'binary'
          });
          workbook.SheetNames.forEach(function(sheetName, idx) {
            var firstSheetLoop = idx == 0 ? true : false;
            var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
            
            var hasExistsAnyUndefined = XL_row_object[0].undefined;

            if(hasExistsAnyUndefined) {
              Toast.fire({
                icon: 'error',
                title: 'Please make sure that all headers are entered in the excel you uploaded.'
              });
              return true;
            }
            stateObject[sheetName] = XL_row_object;
            Toast.fire({
              icon: 'success',
              title: 'Excel successfully converted.'
            });
            firstSheetLoop && loadDefaultHeader(sheetName);
            setSheetNames(sheetName, firstSheetLoop);
            openAccessControllerButtons();
          })
        };

        reader.onerror = function(ex) {
          console.log(ex);
        };

        reader.readAsBinaryString(file);
      };
  };

  function openAccessControllerButtons() {
      document.getElementById('xl__json').classList.remove('disabled');
      document.getElementById('copyofresult').classList.remove('disabled');
      return true;
  }

  function setSheetNames(sheetName, isFirstSheet) {
      var li = document.createElement("li");
      li.setAttribute("class", "nav-item");
      li.innerHTML = `
          <a class="nav-link ${isFirstSheet ? 'active' : ''}" data-value="${sheetName}" href="#">${sheetName}</a>
      `;
      document.getElementById('loadedSheetNames').append(li);
  }

  function changeSheet(e) {
      if(!e.target.classList.contains('active')) {
        $(document).find('#loadedSheetNames .nav-link.active').removeClass('active');
        e.target.classList.add('active');
        clearHeaders();
        var selectedSheet = e.target.dataset.value;
        setHeaders(selectedSheet);
      }
      
  }

  function loadDefaultHeader(sheetName) {
      setHeaders(sheetName);
  }

  function clearHeaders() {
    $(document).find('#resultOfJson').empty();
    return true;
  }

  function setHeaders(sheetName) {
      var item = getSheet(sheetName)[0];

      Object.keys(item).forEach(function(item) {
            var li = document.createElement("li");
            li.setAttribute("class", "nav-item dropdown mb-3");
            li.innerHTML = `
            <a class="nav-link dropdown-toggle added-after" data-bs-toggle="dropdown" href="#" role="button" data-value="${item}" aria-expanded="false"><i class="bi bi-x-diamond-fill"></i> ${item}</a>
                <ul class="dropdown-menu w-100 text-center">
                    <li><h6 class="dropdown-header">Target Rules</h6></li>
                    <li><hr class="dropdown-divider"></li>
                    <li><a class="dropdown-item" data-type="array" data-value="${item}" href="#"><i class="bi bi-arrow-return-right"></i> Extract For Array</a></li>
                    <li><a class="dropdown-item" data-type="rules" data-value="${item}" href="#"><i class="bi bi-arrow-return-right"></i> Extract For Rules</a></li>
                    <li><hr class="dropdown-divider"></li>
                    <li><h6 class="dropdown-header">Widget Rules</h6></li>
                    <li><hr class="dropdown-divider"></li>
                    <li><a class="dropdown-item" data-type="widget-filter" data-value="${item}" href="#"><i class="bi bi-arrow-return-right"></i> Extract For Filter</a></li>
                    <li><a class="dropdown-item" data-type="widget-fixed-product" data-value="${item}" href="#"><i class="bi bi-arrow-return-right"></i> Extract Fixed Product</a></li>
                </ul>
            `;
            document.getElementById('resultOfJson').append(li);
      });
  }

  function clearForNewExcel() {
      $(document).find(".added-after").remove();
      $(document).find('#xlx_html').empty()
      $(document).find('#loadedSheetNames').empty();
      stateObject = {};
      return true;
  }

  function handleFileSelect(evt) {
    var files = evt.target.files; 
    var xl2json = new ExcelToJSON();
    console.log(files[0]);
    xl2json.parseExcel(files[0]);
  }

  function copyAll() {
        var copyText = document.getElementById("xlx_html");
        if(copyText.innerHTML == "") {
          Toast.fire({
            icon: 'error',
            title: 'Please select a data to copy.'
          });
          return true;
        }
        copyText.select();
        document.execCommand("copy");
        Toast.fire({
          icon: 'success',
          title: 'Data successfully copied to clipboard.'
        });
  }

   function getJsonData() {
        var elem = document.getElementById("xlx_html");
        elem.innerHTML = "";
        elem.innerHTML = JSON.stringify(stateObject);
   }

   function getSheet(sheetName) {
      return stateObject[sheetName];
   }

   function setActiveRules(filterName, ruleType) {
        var clickedRule = $(document).find(`#resultOfJson .dropdown-item[data-value='${filterName}'][data-type='${ruleType}']`);
        var clickedParent = $(document).find(`#resultOfJson .added-after[data-value='${filterName}']`);

        if (!clickedParent.hasClass('active')) {
          $(document).find(`#resultOfJson .added-after.active`).removeClass('active');
          clickedParent.addClass('active');
        }

        if (!clickedRule.hasClass('active')) {
          $(document).find(`#resultOfJson .dropdown-item.active`).removeClass('active');
          clickedRule.addClass('active');
        }
   }  

  function handleJSONObject(e) {
    e.preventDefault();
    var elem = document.getElementById("xlx_html");
    elem.innerHTML = "";
    var type = e.target.dataset.type;
    var filterName = e.target.dataset.value;
    var activeSheet = $(document).find("#loadedSheetNames .nav-link.active").attr('data-value');
    setActiveRules(filterName, type);
    var sheet = getSheet(activeSheet);

    Looper(type, sheet, elem, filterName);
  }

  function Looper(type, loopData, element, filterParameter) {
    var parser = null;

    if(type == 'rules') {parser = '\r\n'}
    else if (type == 'widget-filter') parser = ','
    else if (type == 'widget-fixed-product') parser = ';'

    switch(type) {
      case "array":
        var newArr = [];
        loopData.forEach(function(item) {
          newArr.push(item[filterParameter]);
        });
        element.innerHTML = JSON.stringify(newArr);
        break;

      case "rules":
      case "widget-filter":
      case "widget-fixed-product":
        var length = loopData.length;
        var i=0;
        var processedData = "";
        while (length--) {
            processedData += loopData[i][filterParameter] + parser;
            i++;
        }
        element.innerHTML = processedData.slice(0, -1);
    }
  }