'use strict';

// Wrap everything in an anonymous function to avoid polluting the global namespace
(function () {
  $(document).ready(function () {
    tableau.extensions.initializeAsync().then(function () {
      // Since dataSource info is attached to the worksheet, we will perform
      // one async call per worksheet to get every dataSource used in this
      // dashboard.  This demonstrates the use of Promise.all to combine
      // promises together and wait for each of them to resolve.
      let dataSourceFetchPromises = [];

      // Maps dataSource id to dataSource so we can keep track of unique dataSources.
      let dashboardDataSources = {};

      // To get dataSource info, first get the dashboard.
      const dashboard = tableau.extensions.dashboardContent.dashboard;

      // Then loop through each worksheet and get its dataSources, save promise for later.
      dashboard.worksheets.forEach(function (worksheet) {
        dataSourceFetchPromises.push(worksheet.getDataSourcesAsync());
      });

      Promise.all(dataSourceFetchPromises).then(function (fetchResults) {
        fetchResults.forEach(function (dataSourcesForWorksheet) {
          dataSourcesForWorksheet.forEach(function (dataSource) {
            if (!dashboardDataSources[dataSource.id]) { // We've already seen it, skip it.
              dashboardDataSources[dataSource.id] = dataSource;
            }
          });
        });

        $('#refresh-datasources-button').on('click', () => {
          $('#refresh-datasources-button').text('Refreshing...').attr('disabled', true);
          refreshDataSources(dashboardDataSources).then(() => {
            console.log(`Successfully refreshed the datasources.`);
          })
          .catch(reason => console.error(reason))
          .finally(() => {
            $('#refresh-datasources-button').text('Refresh Datasources').attr('disabled', false);
          });
        });
      });

      loadSettings();
    }, function (err) {
      // Something went wrong in initialization.
      console.log('Error while Initializing: ' + err.toString());
    });
  });

  // Refreshes the given dataSources.
  function refreshDataSources (dataSources) {
    let dataSourceRefreshPromises = Object.keys(dataSources).map(datasourceId => dataSources[datasourceId].refreshAsync());

    reloadServerDataset();

    return Promise.all(dataSourceRefreshPromises);
  }

  function loadSettings() {
    loadField('settings-server');
    loadField('settings-email-address');
    loadField('settings-password');
    loadField('settings-pat');
    loadField('settings-site');
    loadField('settings-entity-name');
    loadComboboxField('settings-type', 'settings-type-datasource', 'settings-type-workbook');
  }

  function loadField(id) {
    const field = $(`#${id}`);
    field.val(tableau.extensions.settings.get(id));
    field.change(() => {
      tableau.extensions.settings.set(id, field.val());
      tableau.extensions.settings.saveAsync()
        .then(() => console.log(`saved ${id}`));
    });
    console.log(`loaded ${id}`);
  }

  function loadComboboxField(id) {
    var currentVal = tableau.extensions.settings.get(id);
    for(let i = 1; i < arguments.length; i++) {
      const fieldId = arguments[i];
      const field = $(`#${arguments[i]}`);
      
      field.attr('checked', fieldId == currentVal);

      field.change(() => {
        tableau.extensions.settings.set(id, fieldId);
        tableau.extensions.settings.saveAsync()
          .then(() => console.log(`saved ${id}`));
      });
    }
    console.log(`loaded ${id}`);
  }

  function reloadServerDataset() {
    const server      = tableau.extensions.settings.get('settings-server');
    const email       = tableau.extensions.settings.get('settings-email-address');
    const password    = tableau.extensions.settings.get('settings-password');
    const pat         = tableau.extensions.settings.get('settings-pat');
    const site        = tableau.extensions.settings.get('settings-site');
    const entityName  = tableau.extensions.settings.get('settings-entity-name');
    const isDataset   = tableau.extensions.settings.get('settings-type') == 'settings-type-datasource';
    
    if (!((email && password) || pat || site || entityName)) {
      console.log('Tableau server/online settings are not configured');
      return;
    }

    console.log('Reloading server datasets...');
    tableauApiSignin(server, site, email, password);
  }

  function tableauApiSignin(server, site, email, password) {
    const body = `
      <tsRequest>
        <credentials name="${email}" password="${password}" >
          <site contentUrl="${site}" />
        </credentials>
      </tsRequest>
    `;

    var response = tableauRequest(server, 'api/3.10/auth/signin', 'post', body);
  }

  function tableauRequest(server, subUrl, method, body, sid) {
    const url = server + subUrl;
    const headers = {};
    if (sid) {
      headers['X-Tableau-Auth'] = sid;
    }

    return new Promise((resolve, reject) => {
      $.ajax({
        url: url,
        type: method,
        data: body,
        headers: headers,
      })
      .done((response, textStatus, jqXHR) => {
        console.log(response, textStatus, jqXHR);
      })
      .fail((jqXHR, textStatus, errorThrown) => {
        console.error(jqXHR, textStatus, errorThrown);
      });
    }) 
  }
})();
