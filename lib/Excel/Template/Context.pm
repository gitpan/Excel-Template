package Excel::Template::Context;

use strict;

BEGIN {
    use vars qw(@ISA);
    @ISA = qw(Excel::Template::Base);

    use Excel::Template::Base;
}

use Excel::Template::Format;

# This is a helper object. It is not instantiated by the user, nor does it
# represent an XML object. Rather, every container will use this object to
# maintain the context for its children.

my %isAbsolute = map { $_ => 1 } qw(
    ROW
    COL
);

sub new
{
    my $class = shift;
    my $self = $class->SUPER::new(@_);

    $self->{ACTIVE_WORKSHEET} = undef;
    $self->{ACTIVE_FORMAT}    = Excel::Template::Format->blank_format($self);

    $self->{PARAM_MAP} = [] unless UNIVERSAL::isa($self->{PARAM_MAP}, 'ARRAY');
    $self->{STACK}     = [] unless UNIVERSAL::isa($self->{STACK},     'ARRAY');

    $self->{$_} = 0 for keys %isAbsolute;

    return $self;
}

sub param
{
    my $self = shift;
    my ($param, $depth) = @_;
    $depth ||= 0;

    my $val = undef;
    my $found = 0;

    for my $map (reverse @{$self->{PARAM_MAP}})
    {
        next unless exists $map->{$param};
        $depth--, next if $depth;

        $found = 1;
        $val = $map->{$param};
        last;
    }

    die "Parameter '$param' not found\n"
        if !$found && $self->{DIE_ON_NO_PARAM};

    return $val;
}

sub resolve
{
    my $self = shift;
    my ($obj, $key, $depth) = @_;
    $key = uc $key;
    $depth ||= 0;

    my $obj_val = $obj->{$key};

    $obj_val = $self->param($1)
        if $obj->{$key} =~ /^\$(\S+)$/o;

#GGG Does this adequately test values to make sure they're legal??
    # A value is defined as:
    #    1) An optional operator (+, -, *, or /)
    #    2) A decimal number

#GGG Convert this to use //x
    my ($op, $val) = $obj_val =~ m!^\s*([\+\*\/\-])?\s*([\d.]*\d)\s*$!oi;

    # Unless it's a relative value, we have what we came for.
    return $obj_val unless $op;

    my $prev_val = $isAbsolute{$key}
        ? $self->{$key}
        : $self->get($obj, $key, $depth + 1);

    return $obj_val unless defined $prev_val;
    return $prev_val unless defined $obj_val;

    # Prevent divide-by-zero issues.
    return $val if $op eq '/' and $val == 0;

    my $new_val;
    for ($op)
    {
        /^\+$/ && do { $new_val = ($prev_val + $val); last; };
        /^\-$/ && do { $new_val = ($prev_val - $val); last; };
        /^\*$/ && do { $new_val = ($prev_val * $val); last; };
        /^\/$/ && do { $new_val = ($prev_val / $val); last; };

        die "Unknown operator '$op' in arithmetic resolve\n";
    }

    return $new_val if defined $new_val;
    return;
}

sub enter_scope
{
    my $self = shift;
    my ($obj) = @_;

    push @{$self->{STACK}}, $obj;

    for my $key (keys %isAbsolute)
    {
        next unless exists $obj->{$key};
        $self->{$key} = $self->resolve($obj, $key);
    }

    return 1;
}

sub exit_scope
{
    my $self = shift;
    my ($obj, $no_delta) = @_;

    unless ($no_delta)
    {
        my $deltas = $obj->deltas($self);
        $self->{$_} += $deltas->{$_} for keys %$deltas;
    }

    pop @{$self->{STACK}};

    return 1;
}

sub get
{
    my $self = shift;
    my ($dummy, $key, $depth) = @_;
    $depth ||= 0;
    $key = uc $key;

    return unless @{$self->{STACK}};

    my $obj = $self->{STACK}[-1];

    return $self->{$key} if $isAbsolute{$key};

    my $val = undef;
    my $this_depth = $depth;
    foreach my $e (reverse @{$self->{STACK}})
    {
        next unless exists $e->{$key};
        next if $this_depth-- > 0;

        $val = $self->resolve($e, $key, $depth);
        last;
    }

    $val = $self->{$key} unless defined $val;
    return $val unless defined $val;

    return $self->param($1, $depth) if $val =~ /^\$(\S+)$/o;

    return $val;
}

sub active_format
{
    my $self = shift;
    
    $self->{ACTIVE_FORMAT} = $_[0]
        if @_;

    $self->{ACTIVE_FORMAT};
}

sub active_worksheet
{
    my $self = shift;
    
    $self->{ACTIVE_WORKSHEET} = $_[0]
        if @_;

    $self->{ACTIVE_WORKSHEET};
}

1;
__END__

=head1 NAME

Excel::Template::Context

=head1 PURPOSE

=head1 NODE NAME

=head1 INHERITANCE

=head1 ATTRIBUTES

=head1 CHILDREN

=head1 AFFECTS

=head1 DEPENDENCIES

=head1 USAGE

=head1 AUTHOR

Rob Kinyon (rkinyon@columbus.rr.com)

=head1 SEE ALSO

=cut