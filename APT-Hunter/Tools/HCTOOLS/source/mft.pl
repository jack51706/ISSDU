#! c:\perl\bin\perl.exe
#-----------------------------------------------------------
# Simple $MFT parser 
#   - detects ADSs (prints hex dump if they're resident), and
#     Extended Attributes (may indicate ZeroAccess - 
#     http://journeyintoir.blogspot.com/2012/12/extracting-zeroaccess-from-ntfs.html)
#
#
# To-Do:
#   - Update lookup table creation to account for sequence numbers
#     and identify orphaned files
#
#
# http://msdn.microsoft.com/en-us/library/bb470206%28VS.85%29.aspx
#
# copyright 2014 QAR, LLC
# Author: H. Carvey, keydet89@yahoo.com
#-----------------------------------------------------------
use strict;

my $file = shift || die "You must enter a filename.\n";
die "Could not find $file\n" unless (-e $file);

my $record_sz = getRecordSize($file);

my %tab = buildLookupTable($file,$record_sz);

my %lookup;
foreach my $c (sort {$a <=> $b} keys %tab) {
	$lookup{$c} = getPath($c);
}
undef %tab;

#foreach (sort {$a <=> $b} keys %lookup) {
#	print $_."  ".$lookup{$_}."\n";
#}

# Now that we have a table of paths, we need to run through the MFT
# again, and parse out all of the information; we can replace the 
# record number with the value from the lookup table (ie, %ftable)

parseMFT($file,$record_sz);

#-------------------------------------------------------------
# getRecordSize()
#  
#-------------------------------------------------------------
sub getRecordSize {
	my $file = shift;
	my $sig;
	my $ofs = 0;
	my $data;
	my $sz = 1024;
	
	open(FH,"<",$file);
	binmode(FH);
	seek(FH,$ofs,0);
	read(FH,$data,4);
	
	if ($data eq "FILE") {
		seek(FH,0x1c,0);
		read(FH,$data,4);
		$sz = unpack("V",$data);
	}
	close(FH);
	return $sz;
}

#-------------------------------------------------------------
# cleanStr()
# 'Clean up' Unicode strings; in short, 
#-------------------------------------------------------------
sub cleanStr {
	my $str = shift;
	my @list = split(//,$str);
	my @t;
	my $count = scalar(@list)/2;
	foreach my $i (0..$count) {
		push(@t,$list[$i*2]);
	}
	return join('',@t);
}

#-------------------------------------------------------------
# getTime()
# Translate FILETIME object (2 DWORDS) to Unix time, to be passed
# to gmtime() or localtime()
#-------------------------------------------------------------
sub getTime($$) {
	my $lo = shift;
	my $hi = shift;
	my $t;

	if ($lo == 0 && $hi == 0) {
		$t = 0;
	} else {
		$lo -= 0xd53e8000;
		$hi -= 0x019db1de;
		$t = int($hi*429.4967296 + $lo/1e7);
	};
	$t = 0 if ($t < 0);
	return $t;
}

#-------------------------------------------------------------
# getPath()
# 
#-------------------------------------------------------------
sub getPath {
	my $rec = shift;
	my @path;
	my $root;
	while ($root ne "\.") {
		$root = $tab{$rec}{name};
		push(@path,$root);
		$rec = $tab{$rec}{parent};
	}
	return join('\\',reverse @path);
}

#-------------------------------------------------------------
# buildLookupTable()
# 
#-------------------------------------------------------------
sub buildLookupTable {
	my $file = shift;
	my $sz = shift || 1024;
	my $count = 0;
	my $size = (stat($file))[7];
	my $data;
	my %mft;
	my %lookup;
	
	open(MFT,"<",$file) || die "Could not open $file to read: $!\n";
	binmode(MFT);
	while(($count * $sz) < $size) {
		seek(MFT,$count * $sz,0);
		read(MFT,$data,$sz);
		my %names;
		my $record;
		my $hdr = substr($data,0,42);
		%mft = parseRecordHeader($hdr);
	
		$record = $count;
		$count++;
# record must be in use	
		next unless ($mft{flags} & 0x0001);
		next unless ($mft{sig} eq "FILE");
	
		my $ofs = $mft{attr_ofs};
		my $next = 1;
		while ($next == 1) {
			my $attr = substr($data,$ofs,16);
			my ($type,$len,$res,$name_len,$name_ofs,$flags,$id) = unpack("VVCCvvv",$attr);
			$next = 0 if ($type == 0xffffffff || $type == 0x0000);
# $SIA is always resident, so the extra check doesn't matter
			if ($type == 0x10 && $res == 0) {
# Since we're building a lookup table using record numbers and names, 
# we don't need anything from the $STANDARD_INFORMATION attribute (not at 
# the moment)			
			}
# $FNA is always resident, so the extra check doesn't matter
			elsif ($type == 0x30 && $res == 0) {
				my %fn = parseFNAttr(substr($data,$ofs,$len));
				$names{$fn{name_len}}{name} = $fn{name};
				$names{$fn{name_len}}{parent_ref} = $fn{parent_ref};		
			}
# This is where other attributes would get handled
			else{}		
			$ofs += $len;
		}
# Get the longest name of all $F_N attr in the record	
		my $n = (reverse sort {$a <=> $b} keys %names)[0];
		$lookup{$record}{name} = $names{$n}{name} if ($names{$n}{name} ne "");
		$lookup{$record}{parent} = $names{$n}{parent_ref} if ($names{$n}{name} ne "");
	}
	close(MFT);

	return %lookup;
}

#-------------------------------------------------------------
# 
# 
#-------------------------------------------------------------
sub parseRecordHeader {
	my $hdr = shift;
	my %mft;
# length($data) should be 42 bytes
	$mft{sig} = unpack("A4",substr($hdr,0,4));
	$mft{seq} = unpack("v",substr($hdr,16,2));
	$mft{linkcount} = unpack("v",substr($hdr,18,2));
	$mft{attr_ofs} = unpack("v",substr($hdr,20,2));
	$mft{flags} = unpack("v",substr($hdr,22,2));
	$mft{next_attr_id} = unpack("v",substr($hdr,40,2));	
	return %mft;
}

sub parseSIAttr {
	my $si = shift;
	my %si;
	my ($type,$len,$res,$name_len,$name_ofs,$flags,$id,$sz_content,$ofs_content) 
		= unpack("VVCCvvvVv",substr($si,0,22));
		
	my $content = substr($si,$ofs_content,$sz_content);
	my ($t0,$t1) = unpack("VV",substr($content,0,8));
	$si{c_time} = getTime($t0,$t1);
	my ($t0,$t1) = unpack("VV",substr($content,8,8));
	$si{m_time} = getTime($t0,$t1);
	my ($t0,$t1) = unpack("VV",substr($content,16,8));
	$si{mft_m_time} = getTime($t0,$t1);
	my ($t0,$t1) = unpack("VV",substr($content,24,8));
	$si{a_time} = getTime($t0,$t1);
	$si{flags} = unpack("V",substr($content,32,4));	
		
	return %si;	
}

sub parseFNAttr {
	my $fn = shift;
	my %fn;
	my ($type,$len,$res,$name_len,$name_ofs,$flags,$id,$sz_content,$ofs_content) 
		= unpack("VVCCvvvVv",substr($fn,0,22));
	my $content = substr($fn,$ofs_content,$sz_content);
	$fn{parent_ref} = unpack("V",substr($content,0,4));
	$fn{parent_seq} = unpack("v",substr($content,6,2));
	my ($t0,$t1) = unpack("VV",substr($content,8,8));
	$fn{c_time} = getTime($t0,$t1);
	my ($t0,$t1) = unpack("VV",substr($content,16,8));
	$fn{m_time} = getTime($t0,$t1);
	my ($t0,$t1) = unpack("VV",substr($content,24,8));
	$fn{mft_m_time} = getTime($t0,$t1);
	my ($t0,$t1) = unpack("VV",substr($content,32,8));
	$fn{a_time} = getTime($t0,$t1);
	
	$fn{flags} = unpack("V",substr($content,56,4));
	
	$fn{len_name} = unpack("C",substr($content,64,1));
	$fn{namespace} = unpack("C",substr($content,65,1));

	$fn{len_name} = $fn{len_name} * 2;
#	$fn{len_name} = $fn{len_name} * 2 if ($fn{namespace} > 0);
	$fn{name} = substr($content,66,$fn{len_name});
	$fn{name} = cleanStr($fn{name});
#	$fn{name} = cleanStr($fn{name}) if ($fn{namespace} > 0);
#	$fn{name} =~ s/\00//g;
	$fn{name} =~ s/\x0c/\x2e/g;
	$fn{name} =~ s/[\x01-\x0f]//g;
	return %fn;
}

sub parseMFT {
	my $file = shift;
	my $sz = shift || 1024;
	my %attr_types = (16 => "Standard Information",
                  48 => "File name",
                  64 => "Object ID",
                  128 => "Data",
                  144 => "Index Root",
                  160 => "Index Allocation",
                  176 => "Bitmap");

#-----------------------------------------------------------
# Flags from MFT entry header (use an AND/& operation)
#    00 00 deleted file
#    01 00 allocated file
#    02 00 deleted directory
#    03 00 allocated directory
#-----------------------------------------------------------

	my $count = 0;
	my $size = (stat($file))[7];
	my $data;
	my %mft;

	open(MFT,"<",$file) || die "Could not open $file to read: $!\n";
	binmode(MFT);
	my $t = 1;
	while(($count * $sz) < $size) {
		seek(MFT,$count * $sz,0);
		read(MFT,$data,$sz);
		
		my %mft = parseRecordHeader(substr($data,0,42));
		
#		my $test = unpack("V",substr($data,$mft{attr_ofs},4));
#		$t = 0 if ($test == 0xffffffff);
# using flags, perform an AND operation (ie, &) with flags
#  if ($mft{flags} & 0x0001) - allocated; else unallocated/deleted
#	 if ($mft{flags} & 0x0002) - folder/dir; else file	           
		printf "%-10d %-4s Seq: %-4d Link: %-4d 0x%02x %-4d  Flags: %-4d\n",
	       	$count,$mft{sig},$mft{seq},$mft{linkcount},$mft{attr_ofs},$mft{next_attr_id},$mft{flags};
	  
	  
	  if ($mft{flags} & 0x02) {
	  	print "[FOLDER]\n";
	  }
	  else {
	  	print "[FILE]\n";
	  }
	  
	  print "[DELETED]\n" unless ($mft{flags} & 0x01);
	  
	  if (%lookup) {
	  	if (exists $lookup{$count}) {
	  		print $lookup{$count}."\n";
	  	}
	  }
	  
		$count++;
		next unless ($mft{sig} eq "FILE");
	
		my $ofs = $mft{attr_ofs};
		my $next = 1;
		while ($next == 1) {
			my $attr = substr($data,$ofs,16);
			my ($type,$len,$res,$name_len,$name_ofs,$flags,$id) = unpack("VVCCvvv",$attr);
			$next = 0 if ($type == 0xffffffff || $type == 0x0000);
#			printf "  0x%04x %-4d %-2d  0x%04x 	0x%04x\n",$type,$len,$res,$name_len,$name_ofs unless ($type == 0xffffffff);
# $res == 0 -> data is resident
# $SIA is always resident, so the extra check doesn't matter
			if ($type == 0x10 && $res == 0) {
				my %si = parseSIAttr(substr($data,$ofs,$len));			
				print "    M: ".gmtime($si{m_time})." Z\n";
				print "    A: ".gmtime($si{a_time})." Z\n";
				print "    C: ".gmtime($si{mft_m_time})." Z\n";
				print "    B: ".gmtime($si{c_time})." Z\n";
			}
# $FNA is always resident, so the extra check doesn't matter
			elsif ($type == 0x30 && $res == 0) {
				my %fn = parseFNAttr(substr($data,$ofs,$len));
				print "  FN: ".$fn{name}."  Parent Ref: ".$fn{parent_ref}."  Parent Seq: ".$fn{parent_seq}."\n";
				print "    M: ".gmtime($fn{m_time})." Z\n";
				print "    A: ".gmtime($fn{a_time})." Z\n";
				print "    C: ".gmtime($fn{mft_m_time})." Z\n";
				print "    B: ".gmtime($fn{c_time})." Z\n";
			}
			elsif ($type == 0x80) {
				
				if ($name_len > 0) {
					my $i = substr($data,$ofs,$len);
					my $n = substr($i,$name_ofs,($name_len * 2));
					$n =~ s/\00//g;
					print "**ADS: ".$n."\n";
				}
				
				if ($res == 0) {
					print "[RESIDENT]\n";
				
					print "\n";
					my @d = printData(substr($data,$ofs,$len));
					foreach (0..(scalar(@d) - 1)) {
						print $d[$_]."\n";
					}
					"\n";
				}
				
			}
			elsif ($type == 0xe0) {

				print "**Extended Attribute detected.\n";
			}
# This is where other attributes would get handled
			else{}		
			$ofs += $len;
		}
		print "\n";
#	$count++;
	}
	close(MFT);
}


#-----------------------------------------------------------
# printData()
# subroutine used primarily for debugging; takes an arbitrary
# length of binary data, prints it out in hex editor-style
# format for easy debugging
#-----------------------------------------------------------
sub printData {
	my $data = shift;
	my $len = length($data);
	
	my @display = ();
	
	my $loop = $len/16;
	$loop++ if ($len%16);
	
	foreach my $cnt (0..($loop - 1)) {
# How much is left?
		my $left = $len - ($cnt * 16);
		
		my $n;
		($left < 16) ? ($n = $left) : ($n = 16);

		my $seg = substr($data,$cnt * 16,$n);
		my $lhs = "";
		my $rhs = "";
		foreach my $i ($seg =~ m/./gs) {
# This loop is to process each character at a time.
			$lhs .= sprintf(" %02X",ord($i));
			if ($i =~ m/[ -~]/) {
				$rhs .= $i;
    	}
    	else {
				$rhs .= ".";
     	}
		}
		$display[$cnt] = sprintf("0x%08X  %-50s %s",$cnt,$lhs,$rhs);
	}
	return @display;
}