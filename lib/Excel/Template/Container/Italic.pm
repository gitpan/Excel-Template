package Excel::Template::Container::Italic;

use strict;

BEGIN {
    use vars qw(@ISA);
    @ISA = qw( Excel::Template::Container );

    use Excel::Template::Container;
}

use Excel::Template::Format;

sub render
{
    my $self = shift;
    my ($context) = @_;

    my $old_format = $context->active_format;
    my $format = Excel::Template::Format->copy(
        $context, $old_format,

        italic => 1,
    );
    $context->active_format($format);

    my $child_success = $self->iterate_over_children($context);

    $context->active_format($old_format);
}

1;
__END__

=head1 NAME

Excel::Template::Container::Italic - Excel::Template::Container::Italic

=head1 PURPOSE

To format all children in italic

=head1 NODE NAME

ITALIC

=head1 INHERITANCE

Excel::Template::Container

=head1 ATTRIBUTES

None

=head1 CHILDREN

None

=head1 EFFECTS

None

=head1 DEPENDENCIES

None

=head1 USAGE

<italic>

    ... Children here

</italic>

In the above example, the children will be displayed (if they are displaying
elements) in a italic format. All other formatting will remain the same and the
"italic"-ness will end at the end tag.

=head1 AUTHOR

Rob Kinyon (rkinyon@columbus.rr.com)

=head1 SEE ALSO

BOLD

=cut
