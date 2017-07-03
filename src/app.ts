/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = (reason) => {
    $(document).ready(() => {
      $('#assignRange').click(assignRange);
      $('#copyRange').click(copyRange);
      $('#loadProperty').click(loadProperty);
      $('#singleInputCopy').click(singleInputCopy);
      $('#bindingFromA1Range').click(bindingFromA1Range);
      $('#getAllbdings').click(getAllbdings);
      $('#getBindingData').click(getBindingData);
      $('#getBindingWithOfficeSelect').click(getBindingWithOfficeSelect);
      $('#addHandler').click(addHandler);

      

    });
  };

  async function assignRange() {

    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      console.log('assignRange is clicked');


      var sheet = context.workbook.worksheets.getActiveWorksheet();
      // Values to be updated
      var values = [
        ["Type", "Estimate"],
        ["Transportation", 1670]
      ];
      // Create a proxy object for the range
      var range = sheet.getRange("A1:B2");

      // Assign array value to the proxy object's values property.
      range.values = values;

      // Synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context
      return context.sync().then(function () {
        console.log("Done");
      });
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }

      // await context.sync();
    });

  }


  async function copyRange() {

    console.log('copyRange is clicked');

    // Run a batch operation against the Excel object model. Use the context argument to get access to the Excel document.
    await Excel.run(function (ctx) {

      // Create a proxy object for the range and load the values property
      var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2").load("values");

      // Synchronizes the state between JavaScript proxy objects and real objects in Excel by executing instructions queued on the context
      return ctx.sync().then(function () {
        // Assign the previously loaded values to the new range proxy object. The values will be updated once the following .then() function is invoked.
        ctx.workbook.worksheets.getActiveWorksheet().getRange("B1:B2").values = range.values;
      });
    }).then(function () {
      console.log("done");
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });

  }


  async function loadProperty() {
    console.log('loadProperty is loaded');

    await Excel.run(function (ctx) {
      var sheetName = "Sheet1";
      var rangeAddress = "A1:B2";
      var myRange = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

      myRange.load(["values", "address", "format/*", "format/fill", "entireRow"]);


      return ctx.sync().then(function () {
        console.log(myRange.values);
        console.log(myRange.address); //ok
        console.log(myRange.format.wrapText); //ok
        console.log(myRange.format.fill.color); //ok
        console.log(myRange.format.font.color); //not ok as it was not loaded

      });
    }).then(function () {
      console.log("done");
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }


  async function singleInputCopy() {
    console.log('singleInputCopy is called');

    await Excel.run(function (ctx) {
      var sheetName = 'Sheet1';
      var rangeAddress = 'A1:A20';
      var worksheet = ctx.workbook.worksheets.getItem(sheetName);
      var range = worksheet.getRange(rangeAddress);
      // range.values = 'Due Date';
      range.load('text');
      return ctx.sync().then(function () {
        console.log(range.text);
      });
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });



  }


  async function bindingFromA1Range() {
    Office.context.document.bindings.addFromNamedItemAsync("A1:A3", Office.BindingType.Matrix, { id: "MyCities" },
      function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
          write('Error: ' + asyncResult.error.message);
        }
        else {
          // Write data to the new binding.
          Office.select("bindings#MyCities").setDataAsync([['Berlin'], ['Munich'], ['Duisburg']], { coercionType: "matrix" },
            function (asyncResult) {
              if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                console.log(asyncResult);
                write('Control bound. Binding.id: '
                  + asyncResult.value.id + ' Binding.type: ' + asyncResult.value.type);
              } else
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                  write('Error: ' + asyncResult.error.message);
                }

            });
        }
      });
  }

  function getAllbdings() {
    Office.context.document.bindings.getAllAsync(function (asyncResult) {
      var bindingString = '';
      for (var i in asyncResult.value) {
        bindingString += asyncResult.value[i].id + '\n';
      }
      write('Existing bindings: ' + bindingString);
    });
  }


  function getBindingData() {

    Office.context.document.bindings.getByIdAsync('MyCities', function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
      }
      else {
        write('Retrieved binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);

        let myBinding = asyncResult.value;
        myBinding.getDataAsync(function (asyncResult2) {
          if (asyncResult2.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult2.error.message);
          } else {
            write(asyncResult2.value);
          }
        });
      }
    });

  }

  function getBindingWithOfficeSelect() {
    Office.select("bindings#MyCities", function onError() { }).getDataAsync(function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
      } else {
        write(asyncResult.value);
      }
    }
    )

  }


  function addHandler() {
    Office.select("bindings#MyCities").addHandlerAsync(
      Office.EventType.BindingDataChanged, dataChanged);
  }
  function dataChanged(eventArgs) {
    write('Bound data changed in binding: ' + eventArgs.binding.id);
    getBindingWithOfficeSelect();
  }

  // Function that writes to a div with id='message' on the page.
  function write(message) {
    document.getElementById('message').innerText += message +'\n';
  }




})();
