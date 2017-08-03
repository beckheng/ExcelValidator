#!/usr/bin/perl

#统一的规则,校验通过返回1,失败返回0

use strict;

$|++;

use Spreadsheet::BasicRead;
use File::Find;
use YAML qw(LoadFile);
use Data::Dumper;
use Encode qw(encode decode);

use FindBin qw($Bin);

use lib "$Bin/lib";

use ExcelValidatorUtil;

my $startTime = time();

#处理规则的函数
my %methods = (
	"REQUIRE" => \&_validate_require,
	"MIN" => \&_validate_min,
	"MAX" => \&_validate_max,
	"SUM_VAL" => \&_validate_sum,
	"UNIQUE_WITH" => \&_validate_unique_with,
);

#读取配置参数文件,如果没有则使用和程序目录下一样的config.yml
my $configFile = shift @ARGV;
if (!$configFile){
	$configFile = $Bin . "/config.yml";
}

if (!-e($configFile)){
	die "Config file does not exists: " . $configFile;
}

print "Load config file " . $configFile . "\n";

my $config = LoadFile($configFile);

my $rulePath = $config->{"rulePath"};
my $searchPaths = $config->{"excelPaths"};

foreach my $tmpPath ($rulePath, @$searchPaths){
	# important, because my test env is Windows
	$tmpPath = encode("gbk", $tmpPath);
	
	if (!-e ($tmpPath)){
		die "Path does not exists: " . $tmpPath;
	}
}

#先读入所有校验规则
my %rules = ();
find({ wanted => \&readRules, follow => 0, no_chdir => 0 }, $rulePath);

#遍历要校验的文件
foreach my $p (@$searchPaths){
	print "Start Process " . $p . "\n";
	find({ wanted => \&readExcelFile, follow => 0, no_chdir => 0 }, $p);
	print "End Process " . $p . "\n";
}

#读取所有规则文件
sub readRules{
	my $file = $File::Find::name;
	my $nameOnly = $_;
	
	if (-d $file){
		return;
	}
	
	if ($file !~ /\.xlsx$/i){
		return;
	}
	
	#临时文件也不处理
	if ($file =~ /\/\~/){
		return;
	}
	
	$nameOnly =~ s!\_rules\.xlsx$!\.xlsx!g;
	
	if (exists($rules{$nameOnly})){
		die "Rule for " . $nameOnly . " is duplicate at file " . $file;
	}
	
	print "Load rules file " . $file . "\n";
	
	$rules{$nameOnly} = {};
	
	$file =~ s!\\!\/!g;
	
	my $ss = Spreadsheet::BasicRead->new($file);
	if (!$ss){
		die "Could not open rule file: $file $!";
	}
	
	my $numSheets = $ss->numSheets();
	for (my $sheetIndex = 0; $sheetIndex < $numSheets; $sheetIndex++)
	{
		$ss->setCurrentSheetNum($sheetIndex);
		my $sheetName = $ss->currentSheetName();
		
		if (!exists($rules{$nameOnly}->{$sheetName})){
			$rules{$nameOnly}->{$sheetName} = {};
		}
		
		my %fieldMap = ();
		
		my $data = $ss->getFirstRow();
		for (my $i = 1; $i < scalar(@$data); $i++){
			$fieldMap{uc($data->[$i])} = $i;
		}
		
		while (($data = $ss->getNextRow()))
		{
			if ($ss->getRowNumber() >= 2){
				my $fieldName = $data->[0];
				
				my $fieldValidator = {
					"REQUIRE" => $data->[$fieldMap{"REQUIRE"}],
					"MIN" => $data->[$fieldMap{"MIN"}],
					"MAX" => $data->[$fieldMap{"MAX"}], 
					"SUM_VAL" => $data->[$fieldMap{"SUM_VAL"}],
					"SUM_MIN" => $data->[$fieldMap{"SUM_MIN"}],
					"SUM_MAX" => $data->[$fieldMap{"SUM_MAX"}],
					"SUM_WITH" => $data->[$fieldMap{"SUM_WITH"}],
					"UNIQUE_WITH" => $data->[$fieldMap{"UNIQUE_WITH"}],
				};
				
				$rules{$nameOnly}->{$sheetName}->{$fieldName} = $fieldValidator;
			}
		}
	}
}

sub readExcelFile{
	my $file = $File::Find::name;
	my $nameOnly = $_;
	
	if (-d $file){
		return;
	}
	
	if ($file !~ /\.xlsx$/){
		return;
	}
	
	if (!exists($rules{$nameOnly})){
		return;
	}
	
	my $ss = Spreadsheet::BasicRead->new($file);
	if (!$ss){
		die "Could not open $file $!";
	}
	
	my $curExcelRules = $rules{$nameOnly};
	
	my $numSheets = $ss->numSheets();
	for (my $sheetIndex = 0; $sheetIndex < $numSheets; $sheetIndex++)
	{
		$ss->setCurrentSheetNum($sheetIndex);
		my $sheetName = $ss->currentSheetName();
		
		if (!exists($curExcelRules->{$sheetName})){
			next;
		}
		
		my $curSheetRules = $curExcelRules->{$sheetName};
		
		my ($data);
		while (($data = $ss->getNextRow()))
		{
			if ($ss->getRowNumber() >= $config->{"startRow"}){
				foreach my $col (keys(%$curSheetRules)){
					my $index = ExcelValidatorUtil::colNameToIndex($col);
					my $val = $data->[$index];
					&validate($curSheetRules->{$col}, $val, $data);
				}
			}
		}
	}
}

#统一的校验规则都通过这个函数来调用
sub validate{
	my $rules = shift @_; #要校验的所有规则
	my $val = shift @_; #要校验的数据
	my $data = shift @_; #完整的行数据
	
	foreach my $r (keys %$rules){
		if (!exists($methods{$r})){
			next;
		}
		
		my $m = $methods{$r};
		my $res = $m->($rules, $val, $data);
		if (!$res){
			print "Test Row " . $data->[0] . " for " . $val . " " . $r . " [" . ExcelValidatorUtil::getHumanVal($rules->{$r}) . "] Result " . $res . "\n";
		}
	}
}

sub _validate_require{
	my $rules = shift @_;
	my $val = shift @_;
	my $data = shift @_;
	
	my $testRule = $rules->{"REQUIRE"};
	if ($testRule && (0 == length($val))){
		return 0;
	}
	
	return 1;
}

sub _validate_min{
	my $rules = shift @_;
	my $val = shift @_;
	my $data = shift @_;
	
	my $testRule = $rules->{"MIN"};
	if ((length($testRule) > 0) && ($val < $testRule)){
		return 0;
	}
	
	return 1;
}

sub _validate_max{
	my $rules = shift @_;
	my $val = shift @_;
	my $data = shift @_;
	
	my $testRule = $rules->{"MAX"};
	if ((length($testRule) > 0) && ($val > $testRule)){
		return 0;
	}
	
	return 1;
}

sub _validate_sum{
	my $rules = shift @_;
	my $val = shift @_;
	my $data = shift @_;
	
	my $sumVal = $rules->{"SUM_VAL"};
	my $sumMin = $rules->{"SUM_MIN"};
	my $sumMax = $rules->{"SUM_MAX"};
	my $sumWith = $rules->{"SUM_WITH"};
	
	my @allCellVals = ExcelValidatorUtil::getList($val);
	my @withCells = ExcelValidatorUtil::getList($sumWith);
	foreach my $c (@withCells){
		my $cix = ExcelValidatorUtil::colNameToIndex($c);
		
		push(@allCellVals, ExcelValidatorUtil::getList($data->[$cix]));
	}
	
	my $totalVal = 0;
	foreach my $v (@allCellVals){
		$totalVal += $v;
	}
	
	if ((length($sumVal) > 0) && ($totalVal != $sumVal)){
		return 0;
	}
	
	if ((length($sumMin) > 0) && ($totalVal < $sumMin)){
		return 0;
	}
	
	if ((length($sumMax) > 0) && ($totalVal > $sumMax)){
		return 0;
	}
	
	return 1;
}

sub _validate_unique_with{
	my $rules = shift @_;
	my $val = shift @_;
	my $data = shift @_;
	
	my $uniqueWith = $rules->{"UNIQUE_WITH"};
	if (length($uniqueWith) == 0){
		return 1;
	}
	
	my @allCellVals = ExcelValidatorUtil::getList($val);
	my @withCells = ExcelValidatorUtil::getList($uniqueWith);
	foreach my $c (@withCells){
		my $cix = ExcelValidatorUtil::colNameToIndex($c);
		
		push(@allCellVals, ExcelValidatorUtil::getList($data->[$cix]));
	}
	
	my %toHash = ();
	foreach my $v (@allCellVals){
		$toHash{$v}++;
	}
	
	foreach my $k (keys(%toHash)){
		if ($k > 0 && ($toHash{$k} > 1)){
			return 0;
		}
	}
	
	return 1;
}
