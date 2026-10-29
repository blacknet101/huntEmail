@{
    # The public function names and aliases are the toolkit contract (huntEmail, hunt, ...).
    # Operator status is written to the host on purpose, matching the previous scripts.
    ExcludeRules = @(
        'PSUseApprovedVerbs',
        'PSAvoidUsingWriteHost',
        # UTF-8 without BOM is the repo encoding. Block characters are written as [char] codes.
        'PSUseBOMForUnicodeEncodedFile',
        # -Sender is the existing public parameter name.
        'PSAvoidAssignmentToAutomaticVariable'
    )
}
