package Excel::Template::Container::Bold;

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

        bold => 1,
    );
    $context->active_format($format);

    my $child_success = $self->iterate_over_children($context);

    $context->active_format($old_format);
}

1;
__END__

=head1 NAME

Excel::Template::Container::Bold - Excel::Template::Container::Bold

=head1 PURPOSE

To format all children in bold

=head1 NODE NAME

BOLD

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

<bold>

    ... Children here

</bold>

In the above example, the children will be displayed (if they are displaying
elements) in a bold format. All other formatting will remain the same and the
"bold"-ness will end at the end tag.

=head1 AUTHOR

Rob Kinyon (rkinyon@columbus.rr.com)

=head1 SEE ALSO

Nothing (right now)

=cut
