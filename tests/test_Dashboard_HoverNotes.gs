function test_Dashboard_HoverNotes() {
    // 1. Setup Mocks
    const mockChartBuilder = {
        asColumnChart: function() { return this; },
        addRange: function() { return this; },
        setNumHeaders: function() { return this; },
        setHiddenDimensionStrategy: function() { return this; },
        setOption: function() { return this; },
        setPosition: function() { return this; },
        build: function() { return {}; } // Return a mock chart object
    };

    const mockRange = {
      getValues: function() { return [[]]; },
      getValue: function() { return ''; },
      getNote: function() { return ''; },
      setNote: function() { return this; },
      setValues: function() { return this; },
      setNotes: function() { return this; },
      setNumberFormat: function() { return this; },
      setFontWeight: function() { return this; },
      setBackground: function() { return this; },
      setFontColor: function() { return this; },
      setHorizontalAlignment: function() { return this; },
      setBorder: function() { return this; },
      setValue: function() { return this; },
      applyRowBanding: function() {
        return { setHeaderRowColor: () => ({ setFirstRowColor: () => ({ setSecondRowColor: () => this }) }) };
      },
      offset: function() { return this; },
      clearContent: function() {
        return { clearDataValidations: () => ({ clearNote: () => this }) };
      },
      clear: function() { return this; }
    };

    const mockSheet = {
        getRange: () => mockRange,
        getMaxColumns: () => 20,
        getMaxRows: () => 150,
        getCharts: () => [],
        insertChart: () => {},
        setColumnWidth: () => {},
        hideColumns: () => {},
        deleteSheet: () => {},
        clear: function() { return this; },
        insertRowsAfter: () => {},
        deleteRows: () => {},
        insertColumnsAfter: () => {},
        deleteColumns: () => {},
        newChart: () => mockChartBuilder // Add newChart mock
    };

    const mockSpreadsheet = {
        getSheetByName: (name) => {
            if (name === "Overdue Details") {
                return { ...mockSheet, name: "Overdue Details" };
            }
            return mockSheet;
        },
        deleteSheet: (sheet) => {
            if (sheet.name !== "Overdue Details") {
                throw new Error("Tried to delete the wrong sheet!");
            }
        },
        getUi: () => ({ alert: () => {} })
    };

    global.SpreadsheetApp = {
        getActiveSpreadsheet: () => mockSpreadsheet,
        getUi: () => ({ alert: () => {} }),
        BandingTheme: { LIGHT_GREY: 'LIGHT_GREY' },
        BorderStyle: { SOLID_THIN: 'SOLID_THIN' },
        flush: () => {}, // Add flush mock
        Charts: { ChartHiddenDimensionStrategy: { IGNORE_ROWS: 'IGNORE_ROWS' } } // Add Charts enum
    };

    // Add the missing Session mock directly to the test
    global.Session = {
        getScriptTimeZone: () => 'America/New_York'
    };

    // Spy on the critical functions
    const setNotesSpy = jest.spyOn(mockRange, 'setNotes');
    const deleteSheetSpy = jest.spyOn(mockSpreadsheet, 'deleteSheet');

    // 2. Test Data
    const today = new Date();
    const yesterday = new Date();
    yesterday.setDate(today.getDate() - 1);

    // Helper to create a row with correct column positions based on CONFIG
    const createTestRow = (projectName, deadline, status) => {
        const row = new Array(20).fill(null); // Create a sparse array
        row[CONFIG.FORECASTING_COLS.PROJECT_NAME - 1] = projectName;
        row[CONFIG.FORECASTING_COLS.DEADLINE - 1] = deadline;
        row[CONFIG.FORECASTING_COLS.PROGRESS - 1] = status;
        return row;
    };

    const forecastingData = [
        createTestRow('Project A', yesterday, 'In Progress'), // Overdue
        createTestRow('Project B', today, 'Scheduled'),     // Overdue
        createTestRow('Project C', new Date(today.getTime() + 86400000), 'In Progress'), // Upcoming
        createTestRow('Project D', yesterday, 'Done'), // Not Overdue (terminal status)
    ];

    // Mock readForecastingData to return our test data
    global.readForecastingData = () => ({
        forecastingValues: forecastingData,
        forecastingHeaders: ["Project Name", "Deadline", "Progress"]
    });

    // 3. Run the function
    updateDashboard();

    // 4. Assertions
    const noteCall = setNotesSpy.mock.calls[0][0];
    const generatedNote = noteCall.find(n => n[0] !== null);

    // Assert that only "Project A" is in the overdue note.
    if (!generatedNote || !generatedNote[0].includes("Project A")) {
        throw new Error("Assertion failed: Hover note did not contain the 'In Progress' overdue project.");
    }

    // Assert that "Project B" (Scheduled) is NOT in the overdue note.
    if (generatedNote[0].includes("Project B")) {
        throw new Error("Assertion failed: Hover note incorrectly included a 'Scheduled' project.");
    }

    if (generatedNote[0].includes("Project D")) {
        throw new Error("Assertion failed: Hover note contained a project with a terminal status.");
    }

    if (deleteSheetSpy.mock.calls.length !== 1) {
        throw new Error("Assertion failed: deleteSheet was not called exactly once for 'Overdue Details'.");
    }

    console.log("test_Dashboard_HoverNotes passed.");
}