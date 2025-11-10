# Example usage
class GoogleSheetsManager

  SCOPES = [Google::Apis::SheetsV4::AUTH_SPREADSHEETS].freeze

  def initialize(credentials_path:)
    @credentials_path = credentials_path
    @service = nil
  end

  # Initialize the Google Sheets service with service account credentials
  def service
    @service ||= begin
                   Google::Apis::SheetsV4::SheetsService.new.tap do |s|
                     s.authorization = authorize
                   end
                 end
  end

  def create_spreadsheet(title:, sheet_names: ['Sheet1'])
    spreadsheet = Google::Apis::SheetsV4::Spreadsheet.new(
      properties: Google::Apis::SheetsV4::SpreadsheetProperties.new(title: title),
      sheets: sheet_names.map { |name| create_sheet_properties(name) }
    )

    result = service.create_spreadsheet(spreadsheet)
    puts "Created spreadsheet with ID: #{result.spreadsheet_id}"
    puts "URL: https://docs.google.com/spreadsheets/d/#{result.spreadsheet_id}/edit"
    result
  end

  # Write data to a specific range
  def write_data(spreadsheet_id:, range:, values:)
    value_range = Google::Apis::SheetsV4::ValueRange.new(values: values)
    service.update_spreadsheet_value(
      spreadsheet_id,
      range,
      value_range,
      value_input_option: 'USER_ENTERED'
    )
    puts "Wrote #{values.size} rows to #{range}"
  end

  # Format header row (bold, background color, freeze)
  def format_header_row(spreadsheet_id:, sheet_id: 0)
    requests = [
      # Make first row bold
      {
        repeat_cell: {
          range: {
            sheet_id: sheet_id,
            start_row_index: 0,
            end_row_index: 1
          },
          cell: {
            user_entered_format: {
              text_format: { bold: true },
              background_color: { red: 0.9, green: 0.9, blue: 0.9 }
            }
          },
          fields: 'userEnteredFormat(textFormat, backgroundColor)'
        }
      },
      # Freeze first row
      {
        update_sheet_properties: {
          properties: {
            sheet_id: sheet_id,
            grid_properties: { frozen_row_count: 1 }
          },
          fields: 'gridProperties.frozenRowCount'
        }
      }
    ]

    batch_update(spreadsheet_id, requests)
    puts "Formatted header row"
  end

  # Auto-resize columns to fit content
  def auto_resize_column(spreadsheet_id:, sheet_id: 0, start_column: 0, end_column: 26)
    request = {
      auto_resize_column: {
        dimensions: {
          sheet_id: sheet_id,
          dimension: 'COLUMNS',
          start_index: start_column,
          end_index: end_column
        }
      }
    }

    batch_update(spreadsheet_id, [request])
    puts "Auto-resize columns"
  end

  def add_dropdown_validation(spreadsheet_id:, sheet_id:, range:, values:)
    request = {
      set_data_validation: {
        range: {
          sheet_id: sheet_id,
          start_row_index: range[:start_row],
          end_row_index: range[:end_row],
          start_column_index: range[:start_column],
          end_column_index: range[:end_column]
        },
        rule: {
          condition: {
            type: 'ONE_OF_LIST',
            values: values.map { |v| { user_entered_value: v } }
          },
          show_custom_ui: true,
          strict: true
        }
      }
    }

    batch_update(spreadsheet_id, [request])
    puts "Added dropdown validation"
  end

  # Get spreadsheet metadata
  def get_spreadsheet_info(spreadsheet_id:)
    spreadsheet = service.get_spreadsheet(spreadsheet_id)
    {
      title: spreadsheet.properties.title,
      sheets: spreadsheet.sheets.map do |sheet|
        {
          title: sheet.properties.title,
          sheet_id: sheet.properties.sheet_id,
          row_count: sheet.properties.grid_properties.row_count,
          column_count: sheet.properties.grid_properties.column_count
        }
      end,
      url: spreadsheet.spreadsheet_url(spreadsheet_id)
    }
  end

  private

  def authorize
    Google::Auth::ServiceAccountCredentials.make_creds(
      json_key_io: File.open(@credentials_path),
      scope: SCOPES
    )
  end

  def create_sheet_properties(name)
    Google::Apis::SheetsV4::Sheet.new(
      properties: Google::Apis::SheetsV4::SheetProperties.new(
        title: name,
        grid_properties: Google::Apis::SheetsV4::GridProperties.new(
          row_count: 1000,
          column_count: 26
        )
      )
    )
  end

  def batch_update(spreadsheet_id, requests)
    request_body = Google::Apis::SheetsV4::BatchUpdateSpreadsheetRequest.new(
      requests: requests.map { |r| Google::Apis::SheetsV4::Request.new(r) }
    )
    service.batch_update_spreadsheet(spreadsheet_id, request_body)
  end
end

if __FILE__ == $PROGRAM_NAME
  # Initialize with your service account credentials
  credentials_path = ENV['GOOGLE_CREDENTIALS_PATH'] || 'credentials.json'

  unless File.exist?(credentials_path)
    puts "Error: Credentials file not found at #{credentials_path}"
    puts "Set GOOGLE_CREDENTIALS_PATH environment variable or place credentials.json in current directory"
    exit 1
  end

  manager = GoogleSheetsManager.new(credentials_path: credentials_path)

  # Create a new spreadsheet
  spreadsheet = manager.create_spreadsheet(
    title: 'Test Results Report',
    sheet_names: ['API Tests', 'UI Tests']
  )
  spreadsheet_id = spreadsheet.spreadsheet_id

  # Write header row to API Tests sheet
  headers = [['Test Case', 'Status', 'Duration', 'Error Message', 'Executed By']]
  manager.write_data(
    spreadsheet_id: spreadsheet_id,
    range: 'API Tests!A1:E1',
    values: headers
  )

  # Write sample test data
  test_data = [
    ['Login API Test', 'PASSED', '1.2s', '', 'automation@example.com'],
    ['Get User Data', 'FAILED', '0.8s', 'Timeout error', 'automation@example.com'],
    ['Create Order', 'PASSED', '2.1s', '', 'automation@example.com']
  ]
  manager.write_data(
    spreadsheet_id: spreadsheet_id,
    range: 'API Tests!A2:E4',
    values: test_data
  )

  # Get sheet ID for API Tests (usually 0 for first sheet)
  info = manager.get_spreadsheet_info(spreadsheet_id: spreadsheet_id)
  api_tests_sheet_id = info[:sheets].find { |s| s[:title] == 'API Tests' }[:sheet_id]

  # Format header row
  manager.format_header_row(
    spreadsheet_id: spreadsheet_id,
    sheet_id: api_tests_sheet_id,
  )

  # Add dropdown for Status column (column B, index 1)
  manager.add_dropdown_validation(
    spreadsheet_id: spreadsheet_id,
    sheet_id: api_tests_sheet_id,
    range: { start_row: 1, end_row: 1000, start_column: 1, end_column: 2},
    values: %w[PASSED FAILED SKIPPED BLOCKED]
  )

  # Auto-resize all columns
  manager.auto_resize_column(
    spreadsheet_id: spreadsheet_id,
    sheet_id: api_tests_sheet_id,
    start_column: 0,
    end_column: 5
  )
end