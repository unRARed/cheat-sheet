#!/usr/bin/env ruby

require 'nokogiri'
require 'open-uri'
require 'byebug'
require 'json'
require 'axlsx'
require 'webdrivers'
require 'watir'

Watir.default_timeout = 60
browser = Watir::Browser.new :chrome,
  options: { prefs: {} }, headless: true

if File.file?('injuries.json') && file = File.open("injuries.json").read
  puts 'Using pre-fetched data from injuries.json'
  injury_report = JSON.parse(file, :symbolize_names => true)
else
  puts 'Saving data from ESPN to injuries.json'
  injuries_html = Nokogiri::HTML(
    URI.open('https://www.espn.com/nfl/injuries')
  )
  injuries = JSON.
    parse('{' + injuries_html.css('script').
      map{|s| s.children }[3][0].content.split(/\{/, 2).last[0..-2],
    symbolize_names: true)[:page][:content][:injuries]
  filtered_injuries = injuries.
    map{|team| team[:items].
    select{|item| ['INJURY_STATUS_IR', 'INJURY_STATUS_OUT'].
      include?(item[:type][:name]) &&
      ['QB', 'RB', 'WR', 'TE', 'K'].
      include?(item[:athlete][:position]) }}.
    flatten.
    map{|item| [
      item[:type][:name],
      item[:athlete][:name],
      item[:athlete][:position]
    ]}
  injury_report = {
    ir: filtered_injuries.
      select{|injury| injury[0] == "INJURY_STATUS_IR"},
    out: filtered_injuries.
      select{|injury| injury[0] == "INJURY_STATUS_OUT"}
  }

  File.open('injuries.json', 'w') do |f|
    f.puts injury_report.to_json
  end
end

if File.file?('tiers.json') && file = File.open("tiers.json").read
  puts 'Using pre-fetched data from tiers.json'
  sources = JSON.parse(file, :symbolize_names => true)
else
  puts 'Saving data from Fantasy Pros to tiers.json'
  sources = [
    {
      label: 'qb',
      url: 'https://www.fantasypros.com/nfl/rankings/qb-cheatsheets.php',
      tiers: []
    },
    {
      label: 'rb',
      url: 'https://www.fantasypros.com/nfl/rankings/half-point-ppr-rb-cheatsheets.php',
      tiers: []
    },
    {
      label: 'wr',
      url: 'https://www.fantasypros.com/nfl/rankings/half-point-ppr-wr-cheatsheets.php',
      tiers: []
    },
    {
      label: 'te',
      url: 'https://www.fantasypros.com/nfl/rankings/half-point-ppr-te-cheatsheets.php',
      tiers: []
    },
    {
      label: 'k',
      url: 'https://www.fantasypros.com/nfl/rankings/k-cheatsheets.php',
      tiers: []
    },
    {
      label: 'dst',
      url: 'https://www.fantasypros.com/nfl/rankings/dst-cheatsheets.php',
      tiers: []
    }
  ]
  sources.each do |source|
    doc = Nokogiri::HTML(
      URI.open(source[:url])
    )
    puts "Saving #{source[:label].upcase}s"
    browser.goto(source[:url])
    table = browser.table(id: "ranking-table")
    table.wait_until(&:exists?)
    # match both of:
    #   <tr data-tier="2" class="tier-row static">
    #   <tr class="player-row">
    rows = table.elements(tag_name: "tr", class: /.*-row/)
    rows.wait_until(&:exists?)
    tier = []
    # necessary or will prematurely
    # scrape (missing later tiers)
    sleep 5
    puts "Found #{rows.count} rows"
    rows.each_with_index do |row, index|
      break if index > 120
      if row.attributes[:class].include? "tier-row"
        next if tier.empty?
        puts " -> Scraped Tier #{source[:tiers].count + 1}"
        source[:tiers] << tier
        tier = []
      elsif row.attributes[:class].include? "player-row"
        tier << "#{row.a.text} " \
          "#{row.span(class: "player-cell-team").text}"
      end
    end
  end

  puts "Writing tiers.json"
  File.open('tiers.json', 'w') do |f|
    f.puts sources.to_json
  end
end

puts "Preparing local data for spreadsheet"
###########################################
## Zip the tiers for displaying in a row ##
###########################################
row_contents = []
positions = sources.map{|s| s[:tiers] }
max_tiers = sources.map{|s| s[:tiers].count }.max
max_tiers = 8 if max_tiers > 8
injured_ir = injury_report[:ir].map{|a| "#{a[1]} (#{a[2]})" }.flatten
injured_out = injury_report[:out].map{|a| "#{a[1]} (#{a[2]})" }.flatten
max_tiers.times do |tier_index|
  (
    sources.map{|s| s[:tiers][tier_index] }.
    compact.map{|s| s.count }.max
  ).times do |player_index|
    row_contents << (
      sources.map do |s|
        s[:tiers].dig(tier_index, player_index)&.strip || ''
      end
    ) + ['', injured_ir.shift,  injured_out.shift ]
  end
  row_contents << ["END OF TIER"]
end

#####################
## Build the sheet ##
#####################
puts "Generating cheat-sheet"
Axlsx::Package.new do |p|
  s = p.workbook.styles
  heading = s.add_style fg_color: 'FFFFFF',
    bg_color: '222222', sz: 8, b: true
  normal = s.add_style fg_color: '222222', sz: 6
  divider = s.add_style fg_color: '222222', bg_color: '222222', sz: 1
  qb = s.add_style fg_color: '222222', bg_color: 'ffffd1', sz: 7
  qb2 = s.add_style fg_color: '222222', bg_color: 'f3ffe3', sz: 7
  wr = s.add_style fg_color: '222222', bg_color: 'ecd4ff', sz: 7
  wr2 = s.add_style fg_color: '222222', bg_color: 'dcd3ff', sz: 7
  rb = s.add_style fg_color: '222222', bg_color: 'aff8db', sz: 7
  rb2 = s.add_style fg_color: '222222', bg_color: 'bffcc6', sz: 7
  te = s.add_style fg_color: '222222', bg_color: 'ffccf9', sz: 7
  te2 = s.add_style fg_color: '222222', bg_color: 'fcc2ff', sz: 7
  k = s.add_style fg_color: '222222', bg_color: '85e3ff', sz: 7
  k2 = s.add_style fg_color: '222222', bg_color: 'ace7ff', sz: 7
  dst = s.add_style fg_color: '222222', bg_color: 'ffdf9e', sz: 7
  dst2 = s.add_style fg_color: '222222', bg_color: 'ffdfbf', sz: 7
  inj1 = s.add_style fg_color: '7b0b0b', bg_color: 'f4cdcc', sz: 7, b: true
  inj2 = s.add_style fg_color: '7b0b0b', bg_color: 'edacab', sz: 7
  body = [qb, rb2, wr, te2, k, dst2, divider, inj1, inj2]
  body2 = [qb2, rb, wr2, te, k2, dst, divider, inj1, inj2]
  is_odd = false

  p.workbook.add_worksheet(
    :name => "Cheat Sheet",
    :page_setup => {
      :orientation => :portrait,
      :fit_to_width => 1
    },
    :page_margins => {
      :right => 0.15,
      :left => 0.15,
      :top => 0.15,
      :bottom => 0.15,
    }
  ) do |sheet|
    sheet.add_row sources.map{|s| s[:label].upcase } +
      ['', 'Injured (IR)', 'Injured (OUT)'],
      style: heading, height: 10
    row_contents.each do |row_content|
      if row_content[0] == "END OF TIER"
        is_odd = !is_odd
        sheet.add_row ["", "", "", "", "", "", "", "", ""],
          style: divider,
          height: 3
        next
      end
      if is_odd
        sheet.add_row row_content, style: body, height: 8
      else
        sheet.add_row row_content, style: body2, height: 8
      end
    end
    sheet.column_widths *(sources.count.times.map{ 16 } + [1] + [15, 15])
  end
  p.serialize('cheat-sheet.xlsx')
end
puts "All done. Crack a beer and draft."
