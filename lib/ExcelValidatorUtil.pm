package ExcelValidatorUtil;

#显示可读的值,如ARRAY这些会被转换出来
sub getHumanVal{
	my $val = shift @_;
	
	if (ref($val) eq "ARRAY"){
		return join(",", $val);
	}
	else{
		return $val;
	}
}

#将数组形式的字符串转换为数组结构
sub getList{
	my $val = shift @_;
	
	return split(/[,;]/, $val);
}

#将 A-Z,AA-AZ,AAA,AAZ这样的列名转换为0 base的索引值
sub colNameToIndex{
	my $colName = shift @_;
	
	my $ordA = ord('A');
	
	my $len = length($colName);
	if ($len == 1){
		return ord($colName) - $ordA;
	}
	elsif ($len > 1){
		my @arr = reverse(split(//, $colName)); #反转数组
		
		my $ix = 0;
		for (my $i = 0; $i < scalar(@arr); $i++){
			my $base = ord($arr[$i]) - $ordA;
			
			if (0 == $i){
				$ix += ($base); 
			}
			else{
				$ix += (($base + 1) * (26 ** $i));
			}
		}
		
		return $ix;
	}
	
	return -1;
}

1;
