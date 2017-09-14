use strict;
use warnings;
use Spreadsheet::WriteExcel;
use Win32;
use Win32::API;
use Win32::OLE;
use Cwd;
my $Au3 = Win32::OLE->new("AutoItX3.Control");#Slave AutoItX3 with perl.

my $rootDir= cwd;
my $testDir = "$rootDir\\Test_Cases";

#######################Generate addition test input file.##################################
open my $fileHandle, '>', "$testDir\\addition_test_input_file.txt"
	or die "Cannot create test case file: $testDir\\addition_test_input_file.txt!";

my $numOfTestCases = 5;

while ($numOfTestCases-- > 0) {
    my @numbers = genNumbers();
    addArrayelements($fileHandle, \@numbers); #create a string from numbers array in the form a + b + c... = d, and write to input file.
}
close $fileHandle;
############################################################################################


#Generate text file of expected outputs from addition_test_input_file.txt
open my $_fileHandle,'<',"$testDir\\addition_test_input_file.txt" 
	or die "Cannot create test case file: $testDir\\addition_test_input_file.txt!";

open my $FH, '>', "$testDir\\expected_addition_test_output.txt" 
	or die "Cannot create test case file: $testDir\\expected_addition_test_output.txt!";

while (my $line = <$_fileHandle> ) {
    chomp $line;
    my @values = split /=/, $line;
    my $n = scalar @values;
    print $FH $values[$n - 1];
    print $FH "\n";
}
close $FH;
close $_fileHandle;


open my $input,'<',"$testDir\\addition_test_input_file.txt" 
	or die "Cannot create test case file: $testDir\\addition_test_input_file.txt!";
open my $output, '>',"$testDir\\actual_addition_test_output.txt" 
	or die "Cannot create test case file: $testDir\\actual_addition_test_output.txt!";#open output file for writing test results.

$Au3-> ClipGet();#clear clip board cache.
$Au3-> ClipGet();
$Au3-> ClipGet();
$Au3-> Run("\"$rootDir\\calc.exe\"");#Open calc.exe located in root directory.
$Au3-> Sleep(500);#Give the calculator app enough time to open.

while(my $operation = <$input> ) {
    chomp $operation;

    $Au3-> WinWaitActive("Calculator");
    $Au3-> WinWaitActive("Calculator");#Bring Calculator to focus twice.

    my @operands = split /\+|\=/, $operation;
    pop @operands; #get rid of last element.

    my $res = driveCalculator(\@operands);
    print $output $res;
    print $output "\n";
}

close $input;
close $output;

$Au3-> Sleep(500);#wait 0.5 secs
$Au3-> WinClose("Calculator");#close calculator

#Open actual and expected results files for processing.
open my $actual,'<',"$testDir\\actual_addition_test_output.txt"
	or die "Cannot create test case file: $testDir\\actual_addition_test_output.txt!";#open addition test output filefor writing.
open my $expected, '<',"$testDir\\expected_addition_test_output.txt"
	or die "Cannot create test case file: $testDir\\expected_addition_test_output.txt!";

open my $input_str, '<',"$testDir\\addition_test_input_file.txt"
	or die "Cannot create test case file: $testDir\\addition_test_input_file.txt!";

$/ = undef;#set the input record seperate to undef, so it will read the entire file in as one big input.
my $actual_data = <$actual> ;
my $expected_data = <$expected> ;
my $input_str_data = <$input_str> ;

#transfer all data sets to respective arrays.
my @_actual_data = split /\n/ , $actual_data;
my @_expected_data = split /\n/ , $expected_data;
my @_input_str_data = split /\n/ , $input_str_data;
close $actual;#Close both file handles.
close $expected;
close $input_str;

#export and cross - compare data in excell.
writeToExcel(\@_actual_data, \@_expected_data, \@_input_str_data);

print "Congratulations! All test cases executed successfully!\n";
system("pause");

####################################################################################################################
#SUB-ROUTINE TO CROSS COMPARE DATA FROM ACTUAL_ADDITION_TEST_OUTPUT.TX WITH EXPECTED_ADDITION_TEST_OUTPUT.TXT
#################################################################################################################### 
sub writeToExcel {
    my $_actual_ref = shift;
    my $_expected_ref = shift;
    my $_input_str_ref = shift;

    my @actual_data_arr = @$_actual_ref;
    my @expected_data_arr = @$_expected_ref;
    my @input_str_data_arr = @$_input_str_ref;

    if ((scalar@ actual_data_arr) != (scalar@ expected_data_arr)) {
        print "Aborting....Data file corrupted!\n";
        return;
    }

    my $wb = Spreadsheet::WriteExcel->new("$testDir/results.xls");
    my $ws = $wb-> add_worksheet();
    $ws-> write(0, 0, "Test Case:");
    $ws-> write(0, 1, "Actual Results:");
    $ws-> write(0, 2, "Expected Results:");
    $ws-> write(0, 3, "Test Case Pass/Fail:");

    my $ind1 = 0;
    my $row = 1;

    while ($ind1 < (scalar@ expected_data_arr)) {
        $ws-> write($row, 0, $input_str_data_arr[$ind1]);
        $ws-> write($row, 1, $actual_data_arr[$ind1]);
        $ws-> write($row, 2, $expected_data_arr[$ind1]);
        if ($actual_data_arr[$ind1] == $expected_data_arr[$ind1]) {
            $ws-> write($row, 3, "Pass");
        } else {
            $ws-> write($row, 3, "Fail");
        }
        $row++;
        $ind1++;
    }
}

#############################################################################################
#SUB - ROUTINE TO RUN EACH TEST CASE THROUGH CALCULATOR GUI AND WRITE RESULT TO OUTPUT FILE.
#############################################################################################
sub driveCalculator {
    my $arr_ref = shift;
    my@ numbers = @$arr_ref;
    my $arr_size = scalar@ numbers;

    my $ind = 0;

    while ($ind < $arr_size) {

        my@ digits = split //, $numbers[$ind]; #convert num/string into an array of chars.
        my $_n = scalar @digits;
        for (my $k = 0; $k < $_n; $k++) {#transmit one digit at a time to calculator.
			$Au3->Send("{$digits[$k]}");
			$Au3->Sleep(100);
        }

        $ind++;

        if ($ind < $arr_size) {#if not last input, send a '+' sign to GUI.
			$Au3-> Send("{+}");
        } else {#if reached last input, then hit Enter.
		$Au3-> Send("{=}");
        }
    }

    #extract result from Calculator GUI, wait .5 secs then close.
    $Au3-> Send("{CTRLDOWN}c");#use ctrl + c to capture calculator result and store in clipboard cache.
	$Au3-> Sleep(500);#allow enough timefor the data to be copied to the cache.
	$Au3-> Send("{CTRLUP}");
    $Au3-> ClipPut();
    $Au3-> Sleep(500);#allow enough time for the data to be written to the cache by ClipPut function.
	my $result = $Au3-> ClipGet();
    $Au3-> Sleep(500);#allow enough time for the data to be retrieved from the clipboard cache.
	$Au3-> Send("{ESC}");

    return $result;
}

###############################################################################################################
#TAKE EACH NUMBER IN AN ARRAY AND ADD THEM TOGETHER.PRINT TO TEST CASE FILE IN THE FORM 1 + 2 + 3... = 100.
############################################################################################################### 
sub addArrayelements {
    my $fh = shift;
    my $arr_ref = shift;
    my@ arr = @$arr_ref;
    my $size = scalar@ arr;
    my $sum = 0;

    for (my $i = 0; $i < $size; $i++) {
        print $fh $arr[$i];
        $sum += $arr[$i];
        if ($i == $size - 1) {
            print $fh "=";
        } else {
            print $fh "+";
        }
    }

    print $fh "$sum\n";
}###############################################################################################################



########################################################################
#SUB ROUTINE TO RETURN AN ARRAY OF RANDOMLY GENERATED ARGUMENTS.
########################################################################
sub genNumbers {
    my $random_num = 2 + int(rand(20));#randomly determine numbers to be added.

    my @arr
        = ();
    my $i = 0;

    while ($i < $random_num) {
        my $num = int(rand(10000000));
        $arr[$i] = $num;
        $i++;
    }

    return @arr;
}
#############################################################################