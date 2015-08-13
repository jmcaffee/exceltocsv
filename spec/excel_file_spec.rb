require 'spec_helper'

describe ExcelToCsv::ExcelFile do

  let(:converter) { ExcelToCsv::ExcelFile.new }

  let(:outdir)    { Pathname.new('tmp/spec') }
  let(:testfile)  { Pathname.new('spec/data/test1.xls') }
  let(:outfile)   { outdir + 'test1.csv' }

  before :each do
    outdir.rmtree if outdir.exist? && outdir.directory?
    outdir.mkpath
  end


  it "converts a file to csv" do
    converter.xl_to_csv(testfile, outfile)
    outfile.exist?.should be_truthy
  end

  context "dates" do

    let(:ndatefile)     { Pathname.new('spec/data/normaldate.xls') }
    let(:ndatefileout)  { outdir + 'normaldate.csv' }

    it "normal excel dates are converted consistently" do
      converter.xl_to_csv(ndatefile, ndatefileout)
      ndatefileout.exist?.should be_truthy
      file_to_string(ndatefileout).should include '12/18/2012'
    end


    let(:tdatefile)     { Pathname.new('spec/data/textdate.xls') }
    let(:tdatefileout)  { outdir + 'textdate.csv' }

    it "text excel dates are converted consistently" do
      converter.xl_to_csv(tdatefile, tdatefileout)
      tdatefileout.exist?.should be_truthy
      file_to_string(tdatefileout).should include '12/18/2012'
    end
  end # context "dates"


  context "decimals" do

    let(:decimalsfile)     { Pathname.new('spec/data/decimals.xls') }
    let(:decimalsfileout)  { outdir + 'decimals.csv' }

    it "3 decimal place numbers are processed as is" do
      converter.xl_to_csv(decimalsfile, decimalsfileout)
      decimalsfileout.exist?.should be_truthy
      file_row_starting_with(decimalsfileout, '3 Place Decimal').should include '1.123'
    end

    it "more than 3 decimal places are truncated at 3 places" do
      converter.xl_to_csv(decimalsfile, decimalsfileout)
      decimalsfileout.exist?.should be_truthy
      last_item_from_row('Truncate Decimal', decimalsfileout).should eq '1.234'
    end

    it "decimal places are trucated when 0" do
      converter.xl_to_csv(decimalsfile, decimalsfileout)
      decimalsfileout.exist?.should be_truthy
      last_item_from_row('Integer', decimalsfileout).should eq '1'
    end

    it "trailing decimal zeros are trucated" do
      converter.xl_to_csv(decimalsfile, decimalsfileout)
      decimalsfileout.exist?.should be_truthy
      last_item_from_row('No Trailing Zero', decimalsfileout).should eq '1.23'
    end
  end # context "decimals"


  context "commas" do

    let(:commasfile)     { Pathname.new('spec/data/commastrings.xls') }
    let(:commasfileout)  { outdir + 'commastrings.csv' }

    it "within cells are enclosed in quotes" do
      converter.xl_to_csv(commasfile, commasfileout)
      commasfileout.exist?.should be_truthy
      file_to_string(commasfileout).should include 'Comma String,"This,string,has,commas"'
    end

  end # context "commas"
end # describe ExcelToCsv::ExcelFile
