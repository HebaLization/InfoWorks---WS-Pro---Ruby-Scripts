# ============================================================
# InfoWorks WS Pro
#
# SCENARIO COMPARISON
# Create two Selection Lists:
#
# 1. <Scenario>_New_Objects
#    - New Pipes
#    - New Nodes
#
# 2. <Scenario>_Modified_Objects
#    - Existing Pipes with modified pipe fields
#    - Existing Nodes with modified demand
#
# ============================================================


# ============================================================
# CURRENT NETWORK / DATABASE
# ============================================================

net = WSApplication.current_network
db  = WSApplication.current_database


# ============================================================
# CHECK NETWORK
# ============================================================

if net.nil?

  WSApplication.message_box(
    "No network is currently open.",
    "OK",
    nil,
    false
  )

  exit

end


# ============================================================
# GET SCENARIOS
# ============================================================

scenarios = []

net.scenarios do |scenario|
  scenarios << scenario
end


if scenarios.empty?

  WSApplication.message_box(
    "No scenarios were found.",
    "OK",
    nil,
    false
  )

  exit

end


# ============================================================
# DEFAULT SCENARIOS
# ============================================================

reference_default = scenarios.index("Base")

if reference_default.nil?
  reference_default = 0
end


comparison_default = 0

scenarios.each_with_index do |scenario, index|

  if scenario != "Base"

    comparison_default = index
    break

  end

end


# ============================================================
# USER SELECTS TWO SCENARIOS
# ============================================================

result = WSApplication.prompt(

  "Compare Scenarios - Network Changes",

  [

    [
      "Reference Scenario",
      "String",
      scenarios[reference_default],
      nil,
      "LIST",
      scenarios
    ],

    [
      "Scenario to Compare",
      "String",
      scenarios[comparison_default],
      nil,
      "LIST",
      scenarios
    ]

  ],

  false

)


# ============================================================
# CANCEL
# ============================================================

if result.nil?
  exit
end


reference_scenario  = result[0]
comparison_scenario = result[1]


# ============================================================
# VALIDATE
# ============================================================

if reference_scenario == comparison_scenario

  WSApplication.message_box(
    "Please select two different scenarios.",
    "OK",
    nil,
    false
  )

  exit

end


# ============================================================
# GET SELECTION LIST GROUPS
# ============================================================

selection_groups = []

db.model_object_collection("Selection List Group").each do |group|

  selection_groups << group

end


if selection_groups.empty?

  WSApplication.message_box(
    "No Selection List Group was found.\n\n" +
    "Please create a Selection List Group first.",
    "OK",
    nil,
    false
  )

  exit

end


# ============================================================
# GROUP NAMES
# ============================================================

group_names = []

selection_groups.each do |group|

  group_names << group.name

end


# ============================================================
# SELECT GROUP
# ============================================================

group_result = WSApplication.prompt(

  "Choose Selection List Group",

  [

    [
      "Selection List Group",
      "String",
      group_names[0],
      nil,
      "LIST",
      group_names
    ]

  ],

  false

)


if group_result.nil?
  exit
end


selected_group_name = group_result[0]


# ============================================================
# FIND GROUP
# ============================================================

selection_group = nil

selection_groups.each do |group|

  if group.name == selected_group_name

    selection_group = group
    break

  end

end


if selection_group.nil?

  WSApplication.message_box(
    "Could not find the selected Selection List Group.",
    "OK",
    nil,
    false
  )

  exit

end


# ============================================================
# SELECTION LIST NAMES
# ============================================================

new_list_name =
  "#{comparison_scenario}_New_Objects"


modified_list_name =
  "#{comparison_scenario}_Modified_Objects"


puts ""
puts "================================================"
puts " SCENARIO COMPARISON"
puts "================================================"
puts "Reference : #{reference_scenario}"
puts "Compare   : #{comparison_scenario}"
puts "Group     : #{selected_group_name}"
puts "================================================"


# ============================================================
# ============================================================
# PIPE FIELDS TO COMPARE
# ============================================================
#
# These are meaningful pipe/network attributes.
#
# We deliberately DO NOT compare every field returned by
# field_names because some fields are calculated, internal,
# metadata, results-related, or can produce false differences.
#
# ============================================================

pipe_compare_fields = [

  "us_node_id",
  "ds_node_id",

  "area",

  "local_loss",

  "length",
  "diameter",

  "roughness_type",
  "k",
  "darcy_weissbach",
  "hazen_williams",

  "material",
  "year",

  "bulk_coeff",
  "wall_coeff",

  "system_type",

  "isolation_area",

  "wave_celerity",
  "pressure_rating",

  "bends",

  "iwlive_can_be_closed"

]


# ============================================================
# NODE DEMAND FIELDS TO COMPARE
# ============================================================
#
# wn_node.demand_by_category is a structure.
#
# These are the actual demand fields inside the structure.
#
# ============================================================

node_demand_fields = [

  "category_id",
  "spec_consumption",
  "average_demand",
  "no_of_properties",
  "direct_demand",
  "source",
  "category_type",
  "equivalent_persons"

]


# ============================================================
# HELPER:
# SET CURRENT SCENARIO
# ============================================================

def set_scenario(net, scenario)

  if scenario == "Base"

    net.current_scenario = nil

  else

    net.current_scenario = scenario

  end

end


# ============================================================
# HELPER:
# READ NODE DEMAND
# ============================================================

def node_demand_signature(node, fields)

  result = []


  begin

    demand_structure = node["demand_by_category"]

    unless demand_structure.nil?

      demand_structure.each do |row|

        row_values = []

        fields.each do |field|

          begin

            row_values << row[field]

          rescue

            row_values << nil

          end

        end

        result << row_values

      end

    end

  rescue

    result = []

  end


  # Sort so that category row order does not create a
  # false difference.

  result.sort_by do |row|

    row[0].to_s

  end


  return result

end


# ============================================================
# ============================================================
# READ REFERENCE SCENARIO
# ============================================================
# ============================================================

set_scenario(net, reference_scenario)


puts ""
puts "Reading reference scenario..."
puts reference_scenario


# ------------------------------------------------------------
# REFERENCE PIPES
# ------------------------------------------------------------

reference_pipes = {}


net.row_objects("wn_pipe").each do |pipe|

  values = {}


  pipe_compare_fields.each do |field|

    begin

      values[field] = pipe[field]

    rescue

      values[field] = nil

    end

  end


  reference_pipes[pipe.id] = values

end


puts "Reference pipes: #{reference_pipes.length}"


# ------------------------------------------------------------
# REFERENCE NODES
# ------------------------------------------------------------

reference_nodes = {}


net.row_objects("wn_node").each do |node|

  reference_nodes[node.id] =
    node_demand_signature(
      node,
      node_demand_fields
    )

end


puts "Reference nodes: #{reference_nodes.length}"


# ============================================================
# ============================================================
# READ COMPARISON SCENARIO
# ============================================================
# ============================================================

set_scenario(net, comparison_scenario)


puts ""
puts "Reading comparison scenario..."
puts comparison_scenario


# ============================================================
# RESULT ARRAYS
# ============================================================

new_pipe_ids = []
new_node_ids = []

modified_pipe_ids = []
modified_node_ids = []


# ============================================================
# COUNTERS
# ============================================================

new_pipe_count = 0
new_node_count = 0

modified_pipe_count = 0
modified_node_count = 0

unchanged_pipe_count = 0
unchanged_node_count = 0


# ============================================================
# ============================================================
# COMPARE PIPES
# ============================================================
# ============================================================

puts ""
puts "--------------------------------------------"
puts "Comparing pipes..."
puts "--------------------------------------------"


net.row_objects("wn_pipe").each do |pipe|

  pipe_id = pipe.id


  # ==========================================================
  # NEW PIPE
  # ==========================================================

  unless reference_pipes.has_key?(pipe_id)

    new_pipe_ids << pipe_id

    new_pipe_count += 1

    puts "NEW PIPE      : #{pipe_id}"

    next

  end


  # ==========================================================
  # EXISTING PIPE
  # CHECK ONLY SPECIFIED PIPE FIELDS
  # ==========================================================

  reference_values =
    reference_pipes[pipe_id]

  changed = false

  changed_field = nil


  pipe_compare_fields.each do |field|

    begin

      reference_value =
        reference_values[field]

      comparison_value =
        pipe[field]


      if reference_value != comparison_value

        changed = true
        changed_field = field

        break

      end

    rescue

      # Ignore fields that cannot be compared

    end

  end


  # ==========================================================
  # MODIFIED PIPE
  # ==========================================================

  if changed

    modified_pipe_ids << pipe_id

    modified_pipe_count += 1

    puts "MODIFIED PIPE : #{pipe_id} | #{changed_field}"

  else

    unchanged_pipe_count += 1

  end

end


# ============================================================
# ============================================================
# COMPARE NODES
# ============================================================
# ============================================================

puts ""
puts "--------------------------------------------"
puts "Comparing nodes..."
puts "--------------------------------------------"


net.row_objects("wn_node").each do |node|

  node_id = node.id


  # ==========================================================
  # NEW NODE
  # ==========================================================

  unless reference_nodes.has_key?(node_id)

    new_node_ids << node_id

    new_node_count += 1

    puts "NEW NODE      : #{node_id}"

    next

  end


  # ==========================================================
  # EXISTING NODE
  # COMPARE DEMAND ONLY
  # ==========================================================

  reference_demand =
    reference_nodes[node_id]


  comparison_demand =
    node_demand_signature(
      node,
      node_demand_fields
    )


  # ==========================================================
  # MODIFIED DEMAND
  # ==========================================================

  if reference_demand != comparison_demand

    modified_node_ids << node_id

    modified_node_count += 1

    puts "MODIFIED NODE : #{node_id} | DEMAND"

  else

    unchanged_node_count += 1

  end

end


# ============================================================
# RESULTS
# ============================================================

puts ""
puts "================================================"
puts " COMPARISON RESULTS"
puts "================================================"

puts ""
puts "NEW OBJECTS"
puts "  New Pipes : #{new_pipe_count}"
puts "  New Nodes : #{new_node_count}"

puts ""
puts "MODIFIED OBJECTS"
puts "  Modified Pipes : #{modified_pipe_count}"
puts "  Modified Nodes : #{modified_node_count}"

puts ""
puts "UNCHANGED"
puts "  Pipes : #{unchanged_pipe_count}"
puts "  Nodes : #{unchanged_node_count}"

puts "================================================"


# ============================================================
# FIND EXISTING SELECTION LISTS
# ============================================================

new_list = nil
modified_list = nil


selection_group.children.each do |child|

  if child.name == new_list_name &&
     child.type == "Selection List"

    new_list = child

  end


  if child.name == modified_list_name &&
     child.type == "Selection List"

    modified_list = child

  end

end


# ============================================================
# CREATE NEW OBJECTS LIST
# ============================================================

if new_list.nil?

  begin

    new_list =
      selection_group.new_model_object(
        "Selection List",
        new_list_name
      )

    puts ""
    puts "Created Selection List:"
    puts new_list_name

  rescue => error

    puts ""
    puts "ERROR creating:"
    puts new_list_name
    puts error.message

  end

end


# ============================================================
# CREATE MODIFIED OBJECTS LIST
# ============================================================

if modified_list.nil?

  begin

    modified_list =
      selection_group.new_model_object(
        "Selection List",
        modified_list_name
      )

    puts ""
    puts "Created Selection List:"
    puts modified_list_name

  rescue => error

    puts ""
    puts "ERROR creating:"
    puts modified_list_name
    puts error.message

  end

end


# ============================================================
# ============================================================
# SAVE NEW OBJECTS
# ============================================================
# ============================================================

if new_list != nil

  net.clear_selection


  # ----------------------------------------------------------
  # NEW PIPES
  # ----------------------------------------------------------

  new_pipe_ids.each do |pipe_id|

    pipe =
      net.row_object(
        "wn_pipe",
        pipe_id
      )

    unless pipe.nil?

      pipe.selected = true

    end

  end


  # ----------------------------------------------------------
  # NEW NODES
  # ----------------------------------------------------------

  new_node_ids.each do |node_id|

    node =
      net.row_object(
        "wn_node",
        node_id
      )

    unless node.nil?

      node.selected = true

    end

  end


  # ----------------------------------------------------------
  # SAVE
  # ----------------------------------------------------------

  net.save_selection(new_list)


  puts ""
  puts "Saved NEW OBJECTS:"
  puts "  Pipes: #{new_pipe_count}"
  puts "  Nodes: #{new_node_count}"
  puts "Selection List: #{new_list_name}"

end


# ============================================================
# ============================================================
# SAVE MODIFIED OBJECTS
# ============================================================
# ============================================================

if modified_list != nil

  net.clear_selection


  # ----------------------------------------------------------
  # MODIFIED PIPES
  # ----------------------------------------------------------

  modified_pipe_ids.each do |pipe_id|

    pipe =
      net.row_object(
        "wn_pipe",
        pipe_id
      )

    unless pipe.nil?

      pipe.selected = true

    end

  end


  # ----------------------------------------------------------
  # MODIFIED NODES
  # ----------------------------------------------------------

  modified_node_ids.each do |node_id|

    node =
      net.row_object(
        "wn_node",
        node_id
      )

    unless node.nil?

      node.selected = true

    end

  end


  # ----------------------------------------------------------
  # SAVE
  # ----------------------------------------------------------

  net.save_selection(modified_list)


  puts ""
  puts "Saved MODIFIED OBJECTS:"
  puts "  Pipes: #{modified_pipe_count}"
  puts "  Nodes: #{modified_node_count}"
  puts "Selection List: #{modified_list_name}"

end


# ============================================================
# CLEAR FINAL SELECTION
# ============================================================

net.clear_selection


# ============================================================
# FINAL MESSAGE
# ============================================================

message =

  "Scenario comparison completed.\n\n" +

  "Reference:\n" +
  "#{reference_scenario}\n\n" +

  "Compared:\n" +
  "#{comparison_scenario}\n\n" +

  "NEW OBJECTS\n" +
  "Pipes: #{new_pipe_count}\n" +
  "Nodes: #{new_node_count}\n\n" +

  "MODIFIED OBJECTS\n" +
  "Pipes: #{modified_pipe_count}\n" +
  "Nodes (demand): #{modified_node_count}\n\n" +

  "Selection Lists:\n" +
  "#{new_list_name}\n" +
  "#{modified_list_name}"


WSApplication.message_box(
  message,
  "OK",
  nil,
  false
)


puts ""
puts "================================================"
puts " DONE"
puts "================================================"