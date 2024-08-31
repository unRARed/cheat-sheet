#!/usr/bin/env ruby

require "nokogiri"
require "open-uri"
require "byebug"
require "json"
require "axlsx"
require "webdrivers"
require "watir"

Watir.default_timeout = 60
browser = Watir::Browser.new :firefox, headless: true

options = {
  idp: false,
  concerns: false,
  fresh: false,
  fullppr: false,
}

url_prefix = "half-point-"

ARGV.each do |arg|
  case arg
  when "--idp"
    options[:idp] = true
    puts "Including individual defensive players"
  when "--concerns"
    options[:concerns] = true
    puts "Including list of players having concerns"
  when "--fullppr"
    url_prefix = ""
  when "--fresh"
    File.delete("tiers.json") if File.file?("tiers.json")
  else
    puts "Unknown option: #{arg}"
    exit
  end
end

if options[:concerns]
  STATUSES = ["IR", "IR-R", "NFI-R", "PUP-R", "SUSP"]

  if File.file?("concerns.json") && file = File.open("concerns.json").read
    puts "Using pre-fetched data from concerns.json"
    concerns = JSON.parse(file)
  else
    puts "Saving data from Yahoo to concerns.json"
    concerns_html = Nokogiri::HTML(
      URI.open("https://football.fantasysports.yahoo.com/f1/gamedaycalls")
    )
    rows = concerns_html.css("#gamedayscalltable tbody tr")
    concerns = {}
    rows.each do |row|
      next unless row.css(".ysf-player-name a").text
      # ignore the blank status to keep the list short
      status = row.css("td .Badge-negative-bench").text
      puts "Found #{status}"
      next if status&.strip == "" ||
        !status || !STATUSES.include?(status.strip)

      unless concerns[status]
        concerns[status] = []
      end
      concerns[status] << row.css(".ysf-player-name a").text
    end

    File.open("concerns.json", "w") do |f|
      f.puts concerns.to_json
    end
  end
  flat_concerns = concerns.
    map{|s, players| players.
      map{|p| "#{p} (#{s})" }
    }.flatten.sort
end

if (
  File.file?("tiers.json") &&
  file = File.open("tiers.json").read
)
  puts "Using pre-fetched data from tiers.json"
  sources = JSON.parse(file, :symbolize_names => true)
else
  puts "Saving data from Fantasy Pros to tiers.json"
  sources = [
    {
      label: "qb",
      url: "https://www.fantasypros.com/nfl/rankings/qb-cheatsheets.php",
      tiers: []
    },
    {
      label: "rb",
      url: "https://www.fantasypros.com/nfl/rankings/#{url_prefix}ppr-rb-cheatsheets.php",
      tiers: []
    },
    {
      label: "wr",
      url: "https://www.fantasypros.com/nfl/rankings/#{url_prefix}ppr-wr-cheatsheets.php",
      tiers: []
    },
    {
      label: "te",
      url: "https://www.fantasypros.com/nfl/rankings/#{url_prefix}ppr-te-cheatsheets.php",
      tiers: []
    },
    {
      label: "k",
      url: "https://www.fantasypros.com/nfl/rankings/k-cheatsheets.php",
      tiers: []
    },
    {
      label: "dst",
      url: "https://www.fantasypros.com/nfl/rankings/dst-cheatsheets.php",
      tiers: []
    },
  ]

  if options[:idp]
    sources << {
      label: "idp",
      url: "https://www.fantasypros.com/nfl/rankings/idp-cheatsheets.php",
      tiers: []
    }
  end

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
  File.open("tiers.json", "w") do |f|
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
max_tiers.times do |tier_index|
  (
    sources.map{|s| s[:tiers][tier_index] }.
    compact.map{|s| s.count }.max
  ).times do |player_index|
    row = sources.map do |s|
      s[:tiers].dig(tier_index, player_index)&.strip || ""
    end
    if options[:concerns]
      row += [ flat_concerns.shift ]
    end
    row_contents << row
  end
  row_contents << ["END OF TIER"]
end

#####################
## Build the sheet ##
#####################
puts "Generating cheat-sheet for " \
  "#{sources.map{|s| s[:label].upcase }.join(", ")}"
Axlsx::Package.new do |p|
  s = p.workbook.styles
  heading = s.add_style fg_color: "FFFFFF",
    bg_color: "222222", sz: 8, b: true
  normal = s.add_style fg_color: "222222", sz: 6
  divider = s.add_style fg_color: "222222", bg_color: "222222", sz: 1
  qb = s.add_style fg_color: "222222", bg_color: "ffffd1", sz: 7
  qb2 = s.add_style fg_color: "222222", bg_color: "f3ffe3", sz: 7
  wr = s.add_style fg_color: "222222", bg_color: "ecd4ff", sz: 7
  wr2 = s.add_style fg_color: "222222", bg_color: "dcd3ff", sz: 7
  rb = s.add_style fg_color: "222222", bg_color: "aff8db", sz: 7
  rb2 = s.add_style fg_color: "222222", bg_color: "bffcc6", sz: 7
  te = s.add_style fg_color: "222222", bg_color: "ffccf9", sz: 7
  te2 = s.add_style fg_color: "222222", bg_color: "fcc2ff", sz: 7
  k = s.add_style fg_color: "222222", bg_color: "85e3ff", sz: 7
  k2 = s.add_style fg_color: "222222", bg_color: "ace7ff", sz: 7
  dst = s.add_style fg_color: "222222", bg_color: "ffdf9e", sz: 7
  dst2 = s.add_style fg_color: "222222", bg_color: "ffdfbf", sz: 7
  idp = s.add_style fg_color: "222222", bg_color: "f5f5f5", sz: 7
  idp2 = s.add_style fg_color: "222222", bg_color: "e2e2e2", sz: 7
  inj1 = s.add_style fg_color: "7b0b0b", bg_color: "f4cdcc", sz: 7, b: true
  inj2 = s.add_style fg_color: "7b0b0b", bg_color: "edacab", sz: 7

  body = [qb, rb2, wr, te2, k, dst2]
  body += [idp] if options[:idp]
  body += [inj1] if options[:concerns]

  body2 = [qb2, rb, wr2, te, k2, dst2]
  body2 += [idp] if options[:idp]
  body2 += [inj2] if options[:concerns]

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
    headers = sources.map{|s| s[:label].upcase }
    headers += ["CONCERNS"] if options[:concerns]
    sheet.add_row headers, style: heading, height: 10
    row_contents.each do |row_content|
      if row_content[0] == "END OF TIER"
        is_odd = !is_odd
        sheet.add_row headers.count.times.map { "" },
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
    sheet.column_widths *(headers.count.times.map{ 18 })
  end
  p.serialize("cheat-sheet.xlsx")
end
puts "All done. Crack a beer and draft."
