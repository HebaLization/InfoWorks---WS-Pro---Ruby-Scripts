require 'win32ole'

# Enable verbose debugging
DEBUG = true

net = WSApplication.current_network
db  = WSApplication.current_database

# Helpers
def as_array(coll)
  return coll if coll.is_a?(Array)
  return coll.to_a if coll.respond_to?(:to_a)
  if coll.respond_to?(:each)
    arr = []
    coll.each { |c| arr << c }
    return arr
  end
  []
end

def excel_value_to_string(v)
  return nil if v.nil?
  if v.is_a?(Float)
    v % 1 == 0 ? v.to_i.to_s : v.to_s
  else
    v.to_s
  end
end

# Remove NBSPs, zero-width and normalize whitespace
def sanitize_string(s)
  return nil if s.nil?
  s = excel_value_to_string(s)
  return nil if s.nil?
  # replace non-breaking space and zero-width spaces, normalize whitespace
  s = s.gsub("\u00A0", ' ').gsub(/\p{Cf}/, '').gsub(/\s+/, ' ')
  s
end

def normalize_case_insensitive(s)
  return nil if s.nil?
  sanitize_string(s).downcase
end

def alnum_normalize(s)
  return nil if s.nil?
  sanitize_string(s).gsub(/[^0-9A-Za-z]/, '').downcase
end

def safe_message_box(message, title = "Info")
  begin
    # call the simplest signature to avoid invalid options errors
    WSApplication.message_box(message)
  rescue => e
    puts "#{title}: #{message} (message_box failed: #{e.class}: #{e.message})"
  end
end

def debug_puts(msg)
  puts msg if DEBUG
end

# Collect selection list groups as Ruby array
def collect_selection_list_groups(db)
  groups = []
  queue = as_array(db.root_model_objects)
  until queue.empty?
    working = queue.shift
    groups << working.path if working.type == 'Selection List Group'
    as_array(working.children).each { |c| queue << c }
  end
  groups
end

# Prompt helpers (use simple text prompt for Excel path)
def get_excel_path
  val = WSApplication.prompt("Excel file path (paste full path):", [[ 'Excel File Path', 'String', '' ]], false) rescue nil
  if val.nil? || val.empty? || val[0].to_s.strip == ''
    puts "No Excel file selected - exiting."
    exit
  end
  val[0].to_s
end

def get_selection_group_path(db)
  groups = collect_selection_list_groups(db)
  if groups.empty?
    safe_message_box("No Selection List Group found. Create one first.", "Error")
    exit
  end
  # Use LIST prompt; fall back to text prompt on failure
  begin
    val = WSApplication.prompt("Pick Selection Group", [[ 'Pick Selection Group', 'String', groups.first, 'LIST', groups ]], true)
  rescue => _
    val = WSApplication.prompt("Paste Selection Group path:", [[ 'Selection Group Path', 'String', groups.first ]], true) rescue nil
  end
  if val.nil? || val.empty? || val[0].nil?
    puts "No selection group chosen - exiting."
    exit
  end
  val[0].to_s
end

# Row lookup helper (two-arg API)
def get_row_object(net, id)
  begin
    return net.row_object('_nodes', id)
  rescue => _
  end
  begin
    return net.row_object(id) # try single-arg if supported
  rescue => _
  end
  nil
end

# Read possible Asset ID values from a row using multiple accessors.
# Returns array of candidate strings (sanitized) found for the "Asset ID" field.
def asset_id_candidates_from_row(row)
  return [] if row.nil?
  field = 'Asset ID'
  candidates = []

  begin
    if row.respond_to?(:attribute)
      v = row.attribute(field) rescue nil
      candidates << sanitize_string(v) unless v.nil?
    end
  rescue => _
  end

  begin
    if row.respond_to?(:get)
      v = row.get(field) rescue nil
      candidates << sanitize_string(v) unless v.nil?
    end
  rescue => _
  end

  begin
    if row.respond_to?(:[])
      v = row[field] rescue nil
      candidates << sanitize_string(v) unless v.nil?
    end
  rescue => _
  end

  begin
    if row.respond_to?(:properties)
      props = row.properties rescue nil
      if props.respond_to?(:keys) && (props.key?(field) rescue false)
        v = props[field] rescue nil
        candidates << sanitize_string(v) unless v.nil?
      end
    end
  rescue => _
  end

  begin
    if row.respond_to?(:value)
      v = row.value(field) rescue nil
      candidates << sanitize_string(v) unless v.nil?
    end
  rescue => _
  end

  # Attempt a few common alternative names in case WS Pro stored it under a variant
  ['AssetID', 'Asset_Id', 'ASSET_ID', 'external id', 'ExternalID'].each do |alt|
    begin
      if row.respond_to?(:attribute)
        v = row.attribute(alt) rescue nil
        candidates << sanitize_string(v) unless v.nil?
      end
    rescue => _
    end
    begin
      if row.respond_to?(:[])
        v = row[alt] rescue nil
        candidates << sanitize_string(v) unless v.nil?
      end
    rescue => _
    end
  end

  # Deduplicate and compact
  candidates.map! { |c| c.nil? ? nil : c.to_s }
  candidates.compact!
  candidates.uniq!
  candidates
end

# Attempt multiple comparison strategies between target_id and candidate_id.
# Returns match_type string if matched, else nil.
def compare_id_strategies(target_raw, candidate_raw)
  return nil if target_raw.nil? || candidate_raw.nil?

  t_raw = sanitize_string(target_raw)
  c_raw = sanitize_string(candidate_raw)
  return nil if t_raw.nil? || c_raw.nil?

  # exact (byte-for-byte)
  return 'exact' if t_raw == c_raw

  # trimmed equality (should be same as sanitize_string but keep explicit)
  return 'trim' if t_raw.strip == c_raw.strip

  # case-insensitive
  return 'case_insensitive' if t_raw.downcase == c_raw.downcase

  # alphanumeric-only match (remove punctuation/whitespace)
  return 'alnum' if t_raw.gsub(/[^0-9A-Za-z]/, '').downcase == c_raw.gsub(/[^0-9A-Za-z]/, '').downcase

  nil
end

# Find a row object by matching the "Asset ID" field. Returns [row, match_info]
# match_info is a hash { strategy: '...', candidate: '...', row_ident: '...' } when matched, or nil.
def find_row_by_asset_id(net, id)
  return nil if id.nil? || id.to_s.strip == ''
  target = sanitize_string(id)
  return nil if target.nil? || target == ''

  # Try some direct table lookups first (may succeed if API supports table+id)
  ['_links', '_nodes', '_assets', '_objects'].each do |tbl|
    begin
      ro = net.row_object(tbl, id) rescue nil
      if ro
        debug_puts("Direct table lookup succeeded: table=#{tbl} id=#{id}")
        row_ident = begin; ro.respond_to?(:path) ? ro.path : (ro.respond_to?(:id) ? ro.id.to_s : ro.to_s) rescue ro.to_s; end
        return [ro, { strategy: 'direct_table', candidate: id, row_ident: row_ident }]
      end
    rescue => _
    end
  end

  # Scan accessible collections and compare only "Asset ID"
  collections = []
  begin
    collections << net.row_objects('_links') if net.respond_to?(:row_objects) rescue nil
  rescue => _
  end
  begin
    collections << net.row_objects('_nodes') if net.respond_to?(:row_objects) rescue nil
  rescue => _
  end
  begin
    collections << net.rows if net.respond_to?(:rows) rescue nil
  rescue => _
  end

  collections.compact!
  collections.each do |coll|
    as_array(coll).each do |r|
      begin
        candidates = asset_id_candidates_from_row(r)
        # Print first candidate(s) for debugging
        if DEBUG
          row_ident = begin; r.respond_to?(:path) ? r.path : (r.respond_to?(:id) ? r.id.to_s : r.to_s) rescue r.to_s; end
          debug_puts("Scanning row: #{row_ident} - Asset ID candidates: #{candidates.map{|c| c.to_s[0,120]}.join(' | ')}")
        end
        candidates.each do |cand|
          mt = compare_id_strategies(target, cand)
          if mt
            row_ident = begin; r.respond_to?(:path) ? r.path : (r.respond_to?(:id) ? r.id.to_s : r.to_s) rescue r.to_s; end
            debug_puts("Match found (#{mt}) target='#{target}' candidate='#{cand}' row=#{row_ident}")
            return [r, { strategy: mt, candidate: cand, row_ident: row_ident }]
          end
        end
      rescue => inner
        # ignore row-specific errors but log in debug
        debug_puts("Error scanning row for Asset ID: #{inner.class}: #{inner.message}")
      end
    end
  end

  # Last resort: try the API global lookup by id
  begin
    ro = net.row_object(id) rescue nil
    if ro
      debug_puts("Global row_object(id) lookup succeeded for id=#{id}")
      row_ident = begin; ro.respond_to?(:path) ? ro.path : (ro.respond_to?(:id) ? ro.id.to_s : ro.to_s) rescue ro.to_s; end
      return [ro, { strategy: 'global_lookup', candidate: id, row_ident: row_ident }]
    end
  rescue => _
  end

  nil
end

# MAIN
excel_path = get_excel_path
puts "Using Excel file: #{excel_path}"

sel_group_path = get_selection_group_path(db)
puts "Using Selection List Group path (from prompt): #{sel_group_path}"
mo_selection_group = db.model_object(sel_group_path)

# open excel
excel = WIN32OLE.new('Excel.Application')
begin
  excel.Visible = false
  workbook = excel.Workbooks.Open(excel_path)
  sheet = workbook.Worksheets(1)

  rows = sheet.UsedRange.Rows.Count
  cols = sheet.UsedRange.Columns.Count

  saved_paths = []

  (1..cols).each do |col|
    header_raw = sheet.Cells(1, col).Value
    next if header_raw.nil?
    list_name = header_raw.to_s.strip
    next if list_name.empty?
    list_name = list_name[0,250] if list_name.length > 250

    # Create OR find existing selection list under chosen group
    sel_list = nil
    created = false
    begin
      sel_list = mo_selection_group.new_model_object('Selection List', list_name)
      created = true
    rescue RuntimeError => e
      if e.message =~ /name already in use/i
        # try find among children
        sel_list = as_array(mo_selection_group.children).find { |c| c.type == 'Selection List' && c.name == list_name }
        if sel_list
          created = false
        else
          # fallback to SEL~ path
          begin
            sel_list = db.model_object("#{mo_selection_group.path}>SEL~#{list_name}")
            created = false
          rescue => inner
            raise "Failed to create or find Selection List '#{list_name}': #{inner.class}: #{inner.message}"
          end
        end
      else
        raise
      end
    end

    # Select rows listed in column and save
    net.clear_selection
    missing_ids = []
    matched_by_asset_id = []
    selected_count = 0
    (2..rows).each do |r|
      raw = sheet.Cells(r, col).Value
      next if raw.nil?
      id_raw = excel_value_to_string(raw)
      next if id_raw.nil?
      id = id_raw.strip
      next if id.empty?

      debug_puts("Processing Excel ID: '#{id}'")

      # Try previous direct lookup first
      ro = get_row_object(net, id)
      match_info = nil

      if ro
        debug_puts("Direct get_row_object succeeded for '#{id}'")
      else
        # If not found, try to find by the exact "Asset ID" field
        res = find_row_by_asset_id(net, id)
        if res && res.is_a?(Array)
          ro, match_info = res[0], res[1]
        end
      end

      if ro
        begin
          if ro.respond_to?(:selected=)
            ro.selected = true
          elsif ro.respond_to?(:select)
            ro.select
          else
            # attempt generic assignment if available
            begin
              ro.selected = true
            rescue => inner
              raise inner
            end
          end
          selected_count += 1
          if match_info && match_info[:strategy] != 'direct_table' && match_info[:strategy] != 'global_lookup'
            matched_by_asset_id << "#{id} (#{match_info[:strategy]} cand=#{match_info[:candidate]} row=#{match_info[:row_ident]})"
          elsif match_info
            matched_by_asset_id << "#{id} (#{match_info[:strategy]} row=#{match_info[:row_ident]})"
          end
        rescue => e
          missing_ids << id
          puts "Failed to select row for id '#{id}': #{e.class}: #{e.message}"
        end
      else
        # not found: produce diagnostic of near candidates (scan a limited number of rows and print candidates similar to id)
        diag = []
        begin
          # look for any candidate where alnum normalization matches (helps find hyphen/space issues)
          collections = []
          collections << net.row_objects('_nodes') if net.respond_to?(:row_objects) rescue nil
          collections << net.rows if net.respond_to?(:rows) rescue nil
          collections.compact!
          collections.each do |coll|
            as_array(coll).each do |rr|
              begin
                cands = asset_id_candidates_from_row(rr)
                cands.each do |cand|
                  next if cand.nil?
                  if alnum_normalize(cand) == alnum_normalize(id)
                    ri = begin; rr.respond_to?(:path) ? rr.path : (rr.respond_to?(:id) ? rr.id.to_s : rr.to_s) rescue rr.to_s; end
                    diag << "#{cand[0,120]} (row=#{ri})"
                    break
                  end
                end
                break if diag.size >= 5
              rescue => _
              end
            end
            break if diag.size >= 5
          end
        rescue => _
        end

        missing_ids << id
        puts "Missing Excel ID: '#{id}' - no row found. Near candidates: #{diag.empty? ? 'none found' : diag.join(' | ')}"
      end
    end

    net.save_selection(sel_list)

    # Print what was saved
    path_info = begin
      if sel_list.respond_to?(:path)
        sel_list.path
      else
        "#{mo_selection_group.path}>SEL~#{sel_list.name rescue list_name}"
      end
    rescue => _
      "#{mo_selection_group.path}>#{list_name}"
    end

    puts "Saved Selection List: '#{list_name}' (created?: #{created ? 'yes' : 'no'})"
    puts " -> Selection List object name: #{sel_list.name rescue '<no name available>'}"
    puts " -> Selection List path: #{path_info}"
    puts " -> Selected objects: #{selected_count}; missing: #{missing_ids.size > 0 ? missing_ids.join(', ')[0,200] : 'none'}"
    unless matched_by_asset_id.empty?
      puts " -> Matched by Asset ID: #{matched_by_asset_id.join('; ')}"
    end

    saved_paths << path_info
    net.clear_selection
  end

  # After creation print all children under the chosen group (to help finding them in UI)
  children = as_array(mo_selection_group.children)
  puts "\nChildren of Selection List Group (#{mo_selection_group.path}):"
  children.each do |c|
    begin
      puts " - type: #{c.type.ljust(20)} name: #{c.name}"
    rescue => _
      puts " - <child (unable to read name/type)>"
    end
  end

  # Also print saved paths summary
  puts "\nSummary of saved selection-list paths:"
  saved_paths.each { |p| puts " - #{p}" }

  safe_message_box("Finished creating selection lists. Check script console for saved paths.", "Done")
rescue => e
  safe_message_box("Error: #{e.class}: #{e.message}\nSee script console for details.", "Error")
  puts "Error: #{e.class}: #{e.message}"
  puts e.backtrace.join("\n")
ensure
  begin
    workbook.Close(false) if workbook
    excel.Quit if excel
  rescue => _
  ensure
    workbook = nil
    sheet = nil
    excel = nil
    GC.start
  end
end