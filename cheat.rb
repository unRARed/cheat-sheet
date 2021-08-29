#!/usr/bin/env ruby

require 'nokogiri'
require 'open-uri'
require 'byebug'
require 'json'
require 'axlsx'

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
  puts 'Saving data from borischen.co to tiers.json'
  sources = [
    {
      label: 'qb',
      url: 'http://www.borischen.co/p/quarterback-tier-rankings.html',
      tiers: []
    },
    {
      label: 'rb',
      url: 'http://www.borischen.co/p/half-05-5-ppr-running-back-tier-rankings.html',
      tiers: []
    },
    {
      label: 'wr',
      url: 'http://www.borischen.co/p/half-05-5-ppr-wide-receiver-tier.html',
      tiers: []
    },
    {
      label: 'te',
      url: 'http://www.borischen.co/p/half-05-5-ppr-tight-end-tier-rankings.html',
      tiers: []
    },
    {
      label: 'k',
      url: 'http://www.borischen.co/p/kicker-tier-rankings.html',
      tiers: []
    },
    {
      label: 'dst',
      url: 'http://www.borischen.co/p/defense-dst-tier-rankings.html',
      tiers: []
    }
  ]
  sources.each do |source|
    tiers = []
    doc = Nokogiri::HTML(
      URI.open(source[:url])
    )
    s3_files = doc.css('object').map{|o| o.values[0] }
    s3_files.each do |file|
      puts "Saving #{source[:label]}s"
      text = Net::HTTP.get_response(URI.parse(file))
      text.body.split("\n").each do |line|
        tier = []
        line.split(',').each do |player|
          tier << player.split(/Tier\ \d+:\ /).last
        end
        tiers << tier
      end
    end
    source[:tiers] = tiers
  end

  File.open('data.json', 'w') do |f|
    f.puts sources.to_json
  end
end
###########################################
## Zip the tiers for displaying in a row ##
###########################################
row_contents = []
positions = sources.map{|s| s[:tiers] }
max_tiers = sources.map{|s| s[:tiers].count }.max
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
Axlsx::Package.new do |p|
  puts 'Generating cheat-sheet'
  s = p.workbook.styles
  heading = s.add_style fg_color: 'FFFFFF', bg_color: '222222', sz: 8, b: true
  normal = s.add_style fg_color: '222222', sz: 7
  divider = s.add_style fg_color: '222222', bg_color: '222222', sz: 1
  qb = s.add_style fg_color: '222222', bg_color: 'fbff58', sz: 7
  qb2 = s.add_style fg_color: '222222', bg_color: 'fdff9d', sz: 7
  wr = s.add_style fg_color: '222222', bg_color: '58c8ff', sz: 7
  wr2 = s.add_style fg_color: '222222', bg_color: '9edfff', sz: 7
  rb = s.add_style fg_color: '222222', bg_color: '58ff8c', sz: 7
  rb2 = s.add_style fg_color: '222222', bg_color: '91f1af', sz: 7
  te = s.add_style fg_color: '222222', bg_color: 'ff81ef', sz: 7
  te2 = s.add_style fg_color: '222222', bg_color: 'f5bdee', sz: 7
  k = s.add_style fg_color: '222222', bg_color: 'f7bf55', sz: 7
  k2 = s.add_style fg_color: '222222', bg_color: 'fbd690', sz: 7
  dst = s.add_style fg_color: '222222', bg_color: 'b792f3', sz: 7
  dst2 = s.add_style fg_color: '222222', bg_color: 'd4baff', sz: 7
  inj1 = s.add_style fg_color: '7b0b0b', bg_color: 'e87777', sz: 7, b: true
  inj2 = s.add_style fg_color: '7b0b0b', bg_color: 'ffb1b1', sz: 7
  body = [qb, rb2, wr, te2, k, dst2, divider, inj1, inj2]
  body2 = [qb2, rb, wr2, te, k2, dst, divider, inj1, inj2]
  is_odd = false

  p.workbook.add_worksheet(
    :name => "Cheat Sheet",
    :page_setup => {
      :orientation => :landscape,
      :fit_to_width => 1
    },
    :page_margins => {
      :right => 0.25,
      :left => 0.25,
      :top => 0.25,
      :bottom => 0.25,
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
    sheet.column_widths *(sources.count.times.map{ 15 } + [1] + [16, 16])
  end
  p.serialize('cheat-sheet.xlsx')
end
