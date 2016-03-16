#!/usr/bin/perl -w

=head1 NAME

loginmon.pl - A windows service which records who is logged in

=cut
use threads;
use threads::shared;


use Win32::Daemon;
use Data::Dumper;
use POSIX qw(strftime);
use Win32::OLE('in');
use Socket;
use DBI;
use Pod::Usage;
use Date::Parse;
use Cwd qw/abs_path/;

use YAML;

=head1 SYNOPSIS

loginmon.pl path\loginmon.yaml [-option]

Options:

=over 8

=item B<-inst>

Install the service, initialise the local database

=item B<-uninst>

Uninstall the service

=item B<-initdb>

Initialise the local database (deletes any existing version!)

=item B<-syncdb>

Test the database sync

=back

=cut

my $cfg;

if(!defined($ARGV[0]) or !-r $ARGV[0]) {
	print STDERR "Can't access config file\n";
	pod2usage(1);
	exit
} else {
	$cfg=YAML::LoadFile($ARGV[0]) or die "Config file problem $!";
}

if($#ARGV>=1){
	if($ARGV[1] eq "-inst") { &install }
	elsif($ARGV[1] eq "-uninst") { &uninstall }
	elsif($ARGV[1] eq "-initdb") { &initdb }
	elsif($ARGV[1] eq "-syncdb") { &syncdb ; exit }
	else {
		pod2usage(1);
		exit;
	}
}

open(STDERR,">$cfg->{home}\\log\\".strftime("%Y%m%d-%H%M",localtime).".err.log");
open(STDOUT,">$cfg->{home}\\log\\".strftime("%Y%m%d-%H%M",localtime).".out.log");

my %session :shared = (
	username=>"",
	start_time=>time(), # TODO check this is UTC not local time. Certainly looks like it...
	end_time=>time(),
	id=>0);

my %context =(
	last_state=>SERVICE_STOPPED,
	sqlite=>undef,
	hostname=>"nikon",
	wmi=>undef);

Win32::Daemon::RegisterCallbacks( {
		start => \&cbStart,
		running => \&cbRunning,
		stop => \&cbStop,
		pause => \&cbPause,
		continue => \&cbContinue,
		timer => \&cbRunning,});


Win32::Daemon::StartService(\%context,$cfg->{timeout});

sub uninstall {
	if(Win32::Daemon::DeleteService('','loginmon')) {
		print "Service removed\n";
	} else {
		print "Service not removed. Error info: ".Win32::FormatMessage(Win32::Daemon::GetLastError())."\n";
	}
	exit;
}

sub install {
	my $self=abs_path($0);
	my $conf=abs_path($ARGV[0]);
	my $svcpath=$^X;
	my $svcparm="$self $conf";
  	my %h= (
		machine=>'',
		name=>'loginmon',
		display=>'Log-in time monitor',
		path=>$svcpath,
		user=>'',
		pwd=>'',
		description=>'This service records who is logged in every minute (or timeout configured in config)',
		parameters=>$svcparm,
		dependencies=>"winmgmt");
	#print Dumper(%h);
	if(Win32::Daemon::CreateService(\%h)) {
		print "Service added\n";
	} else {
		print "Service not added. Error info: ".Win32::FormatMessage(Win32::Daemon::GetLastError())."\n";
	}
	&initdb();
	exit;
}

sub cbRunning {
	my ($ev,$context) =@_;
	if(SERVICE_RUNNING == Win32::Daemon::State()) {
		my $user;
		my $machine;
		my $info=$context->{wmi}->ExecQuery("SELECT * FROM Win32_ComputerSystem","WQL", 0x10 || 0x20) or die "WMI Query failed $!" ; # 0x10 return immed, 0x20 forward only
		foreach my $item (in $info) {
			if(defined($item->{Username})) {
				$user = $item->{UserName};
			} else {
				$user = "noone";
			}
			$machine = $item->{Name};
		}
		if($session{'username'} ne $user) {
			# Change of user
			logerr("Session ($session{'id'}) for $session{'username'} from $session{'start_time'} to $session{'end_time'} ended");
			# mark session as ready to sync
			$context->{sqlite}->do("UPDATE sessions SET end_time=?, sent=0 WHERE id=?",undef,$session{'end_time'},$session{'id'});
			$session{'username'}=$user;
			$session{'start_time'}=time();
			$session{'end_time'}=time();
			
			# Update local db 
			$context->{sqlite}->do("INSERT INTO sessions (username,start_time,end_time,sent) VALUES (?,?,?,-1)",undef,$session{'username'},$session{'start_time'},$session{'end_time'});
			$context->{sqlite}->commit();
			# Get a session id from the database
			$session{'id'}=$context->{sqlite}->last_insert_id(undef,undef,'sessions',undef);
			logerr("New session ($session{'id'}) for $session{'username'}");

		} else {
			# Extend session info
			$session{'end_time'}=time();
			# Sync to local db
			$context->{sqlite}->do("UPDATE sessions SET end_time=? WHERE id=?",undef,$session{'end_time'},$session{'id'});
			$context->{sqlite}->commit();
		}
	}
	Win32::Daemon::State(SERVICE_RUNNING);
}

sub cbStart {
	my ($ev,$context)=@_;
	#	dostuff
	$context->{last_state}=SERVICE_RUNNING;
	logerr("Started");
	&dbCon($context);
	&syncdb;
	&wmiCon($context);
	Win32::Daemon::State(SERVICE_RUNNING);
}

sub cbPause {
	my ($ev,$context)=@_;
	$context->{last_state}=SERVICE_PAUSED;
	logerr("Paused");
	Win32::Daemon::State(SERVICE_PAUSED);
}

sub cbContinue {
	my ($ev,$context)=@_;
	$context->{last_state}=SERVICE_RUNNING;
	logerr("Resumed");
	Win32::Daemon::State(SERVICE_RUNNING);
}

sub cbStop {
	my ($ev,$context)=@_;
	$context->{last_state}=SERVICE_STOPPED;
	$context->{sqlite}->do("UPDATE sessions SET end_time=?, sent=0 WHERE id=?",undef,$session{'end_time'},$session{'id'});
	$context->{sqlite}->commit;
	logerr("Stopped");
	Win32::Daemon::State(SERVICE_STOPPED);
	Win32::Daemon::StopService();
}

sub dbCon {
	my($context)=@_;
	$context->{sqlite}=DBI->connect("dbi:SQLite:dbname=$cfg->{home}\\loginmon.sqlite","","",{AutoCommit=>0,RaiseError=>1}) or die "sqlite error";
	# Just in case, mark all unfinished sessions from previous run as ready to sync
	$context->{sqlite}->do("UPDATE sessions SET sent=0 WHERE sent=-1");
	$context->{sqlite}->commit;
}

sub wmiCon {
	my($context)=@_;
	$context->{wmi}=Win32::OLE->GetObject("winmgmts:\\\\localhost\\root\\CIMV2") or die "WMI connection failed";
}

sub logerr {
	my $line=shift(@_);
	print STDERR scalar(localtime).": ".$line."\n";
}

sub initdb {
	unlink("$cfg->{home}\\loginmon.sqlite");
	my $sqlite=DBI->connect("dbi:SQLite:dbname=$cfg->{home}\\loginmon.sqlite","","",{AutoCommit=>0,RaiseError=>1}) or die "sqlite error";
	$sqlite->do("CREATE TABLE sessions (id integer primary key, username text, start_time datetime, end_time datetime, sent integer);") or die "table creation error";
	$sqlite->commit or die "table creation error";
	$sqlite->disconnect;
	print "Database initialised\n";
	exit;
}

sub syncdb {
	&logerr("Starting a sync");
	open(LOG,">$cfg->{home}\\log\\sync_".strftime("%Y%m%d-%H%M",localtime).".log");
	print LOG "Attempting sync\n";
	# connect to local db
	my $sqlite=DBI->connect("dbi:SQLite:dbname=$cfg->{home}\\loginmon.sqlite","","",{AutoCommit=>0,RaiseError=>1}) or die "sqlite error";

	# find all unsynced entries, round start/end times to 15min (900s) intervals
	my $sessions=$sqlite->selectall_hashref("SELECT *,start_time+450-(start_time+450)%900 AS st, end_time+450-(end_time+450)%900 as et FROM sessions WHERE sent=0","id");

	# attempt to connect to remote database
	my $sql=DBI->connect("dbi:$cfg->{remotedb}->{driver}:host=$cfg->{remotedb}->{host};dbname=$cfg->{remotedb}->{name};sslmode=$cfg->{remotedb}->{sslmode}",$cfg->{remotedb}->{user},$cfg->{remotedb}->{pass});
	if(!$sql->ping) {
		print LOG "Remote db connection error ". $sql->errstr. ", aborting sync\n";
		return;
	}

	# Check each local entry
	foreach my $s (sort keys %$sessions) {
		# trim username
		my $username=$sessions->{$s}->{username};
		my $start=$sessions->{$s}->{st};
		my $end=$sessions->{$s}->{et};
		$username=~ s/\S+\\//;
		print(LOG "$username: $start-$end\n");
		if($start != $end and $username ne "noone") {
			# Find if we overlap
		        print(LOG "Checking for clashes\n");
			my $clashes=$sql->selectall_hashref("SELECT * FROM mrbs_entry WHERE start_time <= ? AND end_time >= ? AND room_id=?","id",undef,
				$end,$start,$cfg->{room_id});
			# truncate, delete or split the originals
			foreach my $c (keys %$clashes) {
		                print(LOG "CL: $clashes->{$c}->{start_time}-$clashes->{$c}->{end_time}\n");
				if($clashes->{$c}->{start_time} >= $start && $clashes->{$c}->{end_time} <= $end) {
					# Totally eclipsed
					print(LOG "CL: totally eclipsed\n");
					$sql->do("DELETE FROM mrbs_entry WHERE id=?",undef,$c);
				} elsif($clashes->{$c}->{start_time} < $start && $clashes->{$c}->{end_time} <= $end) {
					# Start kept, end clipped
					print(LOG "CL: end clipped\n");
					$sql->do("UPDATE mrbs_entry SET end_time=?,modified_by=? WHERE id=?",undef,$start,"loginmon",$c);
				} elsif($clashes->{$c}->{start_time} >= $start && $clashes->{$c}->{end_time} > $end) {
					# Start clipped, end kept
					print(LOG "CL: start clipped\n");
					$sql->do("UPDATE mrbs_entry SET start_time=?,modified_by=? WHERE id=?",undef,$end,"loginmon",$c);
				} else {
					# middle snipped. Truncate start
					print(LOG "CL: Middle. Changing start\n");
					$sql->do("UPDATE mrbs_entry SET end_time=?,modified_by=? WHERE id=?",undef,$start,"loginmon",$c);
					# Add in a new end
					print(LOG "CL: Middle. Added new ender\n");
					my @fields=grep(!/^id$/,keys(%{$clashes->{$c}}));
					my $fieldlist=join(", ",@fields);
					my $field_ph=join(", ", map('?',@fields));
					$clashes->{$c}->{modified_by}="loginmon";
					$clashes->{$c}->{start_time}=$end;
					$sql->do("INSERT INTO mrbs_entry ($fieldlist) VALUES ($field_ph)",undef,@{$clashes->{$c}}{@fields});
				}

			}
			# insert
			if(!$sql->do("INSERT INTO mrbs_entry (start_time,end_time,room_id,create_by,modified_by,name,type,description,ical_uid) VALUES (?,?,?,?,?,?,?,?,'MRBS-'||SUBSTRING(MD5(CAST(RANDOM() AS varchar(255))) FROM 1 FOR 20)||'\@websvc')",undef,$start,$end,$cfg->{room_id},$username,"loginmon",$username,'E','Used') ) { 
				print LOG "Insert failed $? $sql->errstr giving up" ;
			       	return;
		       	};
			# TODO: check?
		} else {
			print(LOG "Not sending (short or noone)\n");
		}
		# set flag in local db
		$sqlite->do("UPDATE sessions SET sent=1 WHERE id=?",undef,$s);
		$sqlite->commit;
	}
	$sqlite->disconnect();

	print LOG "Successfully reached end of sync process!\n";
	close(LOG);
	&logerr("Finished a sync");
}

