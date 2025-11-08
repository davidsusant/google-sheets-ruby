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
end