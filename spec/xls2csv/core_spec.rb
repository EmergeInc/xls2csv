# coding:utf-8
require 'spec_helper'

module Xls2Csv
  describe Core do
    let(:output_path){ 'spec/fixture/output' }

    describe '【正常系】xlsを入力する' do
      let(:input_path){ 'spec/fixture/test.xls' }
      let(:sheets){ %w[Sheet1 Sheet2 Sheet3] }
      let(:core){ Core.new(input_path, output_path) }

      it '#initialize' do
        core.class.should == Xls2Csv::Core
      end

      describe '#read' do
        let(:read){ core.read_xls }

        it 'Hashを返す' do
          read.class.should == Hash
        end

        it 'Hashにxlsのシートが格納されている' do
          read.keys.should == sheets
        end

        describe '#write' do
          it '配列をcsvファイルに書き込む' do
            core.write_csv(read.keys.first, read.values.first)
            check_hello_world read.keys.first
          end

          it 'Excelの関数は計算結果をcsvに書き込む' do
            core.write_csv(read.keys[1], read.values[1])
            check_hello_world read.keys.first
          end
        end
      end

      describe '#convert' do
        it 'convertを実行すると全部自動でやってくれる' do
          core.convert
          sheets.each { |s| check_hello_world s }
        end
      end

      def check_hello_world(sheet)
        head = open("#{output_path}/#{sheet}.csv").first
        head.should == "\"hello\",\"world\"\n"
      end
    end

    describe '【異常系】' do
      it '存在しないファイルを入力する' do
        core = Core.new('hoge'*5 + '.xls', output_path)
        core.should be_nil
      end

      it 'xlsでないファイルを入力する' do
        pending '(´・ω・｀)'
      end
    end
  end
end
