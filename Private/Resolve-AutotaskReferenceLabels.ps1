<#
.SYNOPSIS
    Resolves Autotask entity reference IDs to human-readable names.

.DESCRIPTION
    Examines the field metadata for an Autotask entity and identifies fields where isReference is true.

    The referenced IDs are grouped by target entity and queried in batches.
    Resolved names are then added to the original returned objects.

    Examples:

        contractID             -> contractName
        companyID              -> companyName
        assignedResourceID     -> assignedResourceName
        assignedResourceRoleID -> assignedResourceRoleName

    Reference metadata and resolved entity labels are cached for the current PowerShell session to avoid unnecessary repeat API calls.

    This helper resolves entity references only. Picklist fields should continue to be resolved through Get-AutotaskPicklistMeta.

.PARAMETER Resource
    The Autotask entity type of the supplied objects, such as Tickets, Tasks, Contracts or ConfigurationItems.

.PARAMETER InputObject
    One or more returned Autotask entity objects.

.PARAMETER BatchSize
    Maximum number of referenced IDs included in each Autotask query.

.PARAMETER ExcludeField
    One or more reference fields that should not be resolved.

.PARAMETER OverwriteExisting
    Allows an existing resolved-name property to be overwritten.

.EXAMPLE
    $returnedItems = @(
        Resolve-AutotaskReferenceLabels -Resource Tickets -InputObject $returnedItems
    )

.EXAMPLE
    $returnedItems = @(
        Resolve-AutotaskReferenceLabels -Resource Tickets -InputObject $returnedItems -ExcludeField creatorResourceID
    )

.OUTPUTS
    The original Autotask objects with resolved name properties added.

.NOTES
    Intended for use by Get-AutotaskAPIResource when -ResolveAllLabels is set.
#>
function Resolve-AutotaskReferenceLabels {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Resource,

        [Parameter(Mandatory, ValueFromPipeline)]
        [AllowNull()]
        [object[]]$InputObject,

        [ValidateRange(1, 500)]
        [int]$BatchSize = 200,

        [string[]]$ExcludeField = @(),

        [switch]$OverwriteExisting
    )

    begin {
        $inputBuffer = New-Object 'System.Collections.Generic.List[object]'

        if (-not $Script:AutotaskReferenceMetaCache) {
            $Script:AutotaskReferenceMetaCache = @{}
        }

        if (-not $Script:AutotaskReferenceLabelCache) {
            $Script:AutotaskReferenceLabelCache = @{}
        }
    }

    process {
        foreach ($object in @($InputObject)) {
            if ($null -ne $object) {
                [void]$inputBuffer.Add($object)
            }
        }
    }

    end {
        $returnedItems = @($inputBuffer.ToArray())

        if ($returnedItems.Count -eq 0) {
            return
        }

        if (
            -not $Script:AutotaskAuthHeader -or
            -not $Script:AutotaskBaseURI
        ) {
            throw 'ERROR: Autotask API auth is not initialised. Run Add-AutotaskAPIAuth first.'
        }

        $baseURIKey = $Script:AutotaskBaseURI.TrimEnd(
            '/'
        ).ToLowerInvariant()

        $resourceKey = $Resource.Trim().ToLowerInvariant()

        $metadataCacheKey = '{0}|{1}' -f (
            $baseURIKey
        ), (
            $resourceKey
        )

        # Converts the singular referenceEntityType returned by Autotask field metadata into the REST API resource name used in URLs
        $getReferenceResource = {
            param(
                [Parameter(Mandatory)]
                [string]$ReferenceEntityType
                )
                $entityType = $ReferenceEntityType.Trim()
                $entityKey  = $entityType.ToLowerInvariant()

    # Explicit mappings make known Autotask entity names predictable
    $resourceNameMap = @{
        'company'                   = 'Companies'
        'companylocation'           = 'CompanyLocations'
        'configurationitemcategory' = 'ConfigurationItemCategories'
        'contact'                   = 'Contacts'
        'contract'                  = 'Contracts'
        'department'                = 'Departments'
        'opportunity'               = 'Opportunities'
        'product'                   = 'Products'
        'project'                   = 'Projects'
        'resource'                  = 'Resources'
        'role'                      = 'Roles'
        'service'                   = 'Services'
        'servicebundle'             = 'ServiceBundles'
        'task'                      = 'Tasks'
        'ticket'                    = 'Tickets'
    }

    if ($resourceNameMap.ContainsKey($entityKey)) {
        return $resourceNameMap[$entityKey]
    }

    # Generic fallback for reference entity types not explicitly mapped
    if ($entityType -match '(?i)[^aeiou]y$') {
        return $entityType.Substring(
            0,
            $entityType.Length - 1
        ) + 'ies'
    }

    if ($entityType -match '(?i)(s|x|z|ch|sh)$') {
        return $entityType + 'es'
    }

    return $entityType + 's'
}

        # Local property accessor. This avoids errors under strict mode when an entity does not contain a candidate display property
        $getPropertyValue = {
            param(
                [object]$Object,
                [string]$PropertyName
            )

            if ($null -eq $Object) {
                return $null
            }

            $property = $Object.PSObject.Properties[$PropertyName]

            if ($null -ne $property) {
                return $property.Value
            }

            return $null
        }

        # Determines the most useful human-readable label for an entity
        $getReferenceLabel = {
            param(
                [string]$ReferenceEntityType,
                [object]$ResolvedObject
            )

            $entityKey = $ReferenceEntityType.Trim().ToLowerInvariant()

            switch ($entityKey) {
                'resources' {
                    $nameParts = @(
                        & $getPropertyValue $ResolvedObject 'firstName'
                        & $getPropertyValue $ResolvedObject 'lastName'
                    ) | Where-Object {
                        -not [string]::IsNullOrWhiteSpace([string]$_)
                    }

                    if ($nameParts.Count -gt 0) {
                        return ($nameParts -join ' ')
                    }
                }

                'contacts' {
                    $nameParts = @(
                        & $getPropertyValue $ResolvedObject 'firstName'
                        & $getPropertyValue $ResolvedObject 'lastName'
                    ) | Where-Object {
                        -not [string]::IsNullOrWhiteSpace([string]$_)
                    }

                    if ($nameParts.Count -gt 0) {
                        return ($nameParts -join ' ')
                    }
                }

                'tickets' {
                    $ticketNumber = & $getPropertyValue `
                        $ResolvedObject `
                        'ticketNumber'

                    $title = & $getPropertyValue `
                        $ResolvedObject `
                        'title'

                    if (
                        -not [string]::IsNullOrWhiteSpace(
                            [string]$ticketNumber
                        ) -and
                        -not [string]::IsNullOrWhiteSpace(
                            [string]$title
                        )
                    ) {
                        return "$ticketNumber - $title"
                    }

                    if (
                        -not [string]::IsNullOrWhiteSpace(
                            [string]$ticketNumber
                        )
                    ) {
                        return [string]$ticketNumber
                    }

                    if (
                        -not [string]::IsNullOrWhiteSpace(
                            [string]$title
                        )
                    ) {
                        return [string]$title
                    }
                }

                'tasks' {
                    $taskNumber = & $getPropertyValue `
                        $ResolvedObject `
                        'taskNumber'

                    $title = & $getPropertyValue `
                        $ResolvedObject `
                        'title'

                    if (
                        -not [string]::IsNullOrWhiteSpace(
                            [string]$taskNumber
                        ) -and
                        -not [string]::IsNullOrWhiteSpace(
                            [string]$title
                        )
                    ) {
                        return "$taskNumber - $title"
                    }

                    if (
                        -not [string]::IsNullOrWhiteSpace(
                            [string]$taskNumber
                        )
                    ) {
                        return [string]$taskNumber
                    }

                    if (
                        -not [string]::IsNullOrWhiteSpace(
                            [string]$title
                        )
                    ) {
                        return [string]$title
                    }
                }
            }

            # Entity-specific display properties
            $preferredProperties = @{
                billingcodes       = @('name')
                companies          = @('companyName')
                companylocations   = @('name')
                configurationitems = @(
                    'referenceTitle'
                    'serialNumber'
                )
                contracts          = @('contractName')
                departments        = @('name')
                opportunities      = @('title')
                products           = @(
                    'name'
                    'productName'
                )
                projects           = @('projectName')
                roles              = @('name')
                servicebundles     = @(
                    'name'
                    'serviceBundleName'
                )
                services           = @(
                    'name'
                    'serviceName'
                )
            }

            $candidateProperties = @()

            if ($preferredProperties.ContainsKey($entityKey)) {
                $candidateProperties += @(
                    $preferredProperties[$entityKey]
                )
            }

            # Generic fallback properties for entity types not explicitly included above
            $candidateProperties += @(
                'companyName'
                'contractName'
                'projectName'
                'productName'
                'serviceName'
                'serviceBundleName'
                'referenceTitle'
                'ticketNumber'
                'taskNumber'
                'name'
                'title'
                'description'
                'userName'
                'emailAddress'
                'email'
            )

            foreach (
                $propertyName in @(
                    $candidateProperties |
                        Select-Object -Unique
                )
            ) {
                $value = & $getPropertyValue `
                    $ResolvedObject `
                    $propertyName

                if (
                    -not [string]::IsNullOrWhiteSpace(
                        [string]$value
                    )
                ) {
                    return [string]$value
                }
            }

            return $null
        }

        # Retrieve and cache reference field metadata for the source entity
        if (
            $Script:AutotaskReferenceMetaCache.ContainsKey(
                $metadataCacheKey
            )
        ) {
            Write-Information "INFO: Cached Reference Field index hit for resource '$Resource'."

            $referenceFields = @(
                $Script:AutotaskReferenceMetaCache[
                    $metadataCacheKey
                ]
            )
        }
        else {
            $fieldsURI = '{0}/V1.0/{1}/entityInformation/fields' -f (
                $Script:AutotaskBaseURI.TrimEnd('/')
            ), (
                $Resource
            )

            Write-Information "INFO: Building initial Reference Field index for '$Resource' from: $fieldsURI..."

            try {
                $fieldInformation = Invoke-RestMethod `
                    -Method Get `
                    -Uri $fieldsURI `
                    -Headers $Script:AutotaskAuthHeader

                $referenceFields = @(
                    $fieldInformation.fields |
                        Where-Object {
                            $_.isReference -eq $true -and
                            -not [string]::IsNullOrWhiteSpace(
                                [string]$_.referenceEntityType
                            ) -and
                            $_.name -notin $ExcludeField
                        } |
                        ForEach-Object {
                            $outputPropertyName = if (
                                $_.name -match '(?i)ID$'
                            ) {
                                $_.name.Substring(
                                    0,
                                    $_.name.Length - 2
                                ) + 'Name'
                            }
                            else {
                                "$($_.name)Name"
                            }

                            [pscustomobject]@{
                                FieldName           = [string]$_.name
                                ReferenceEntityType = [string]$_.referenceEntityType
                                OutputPropertyName  = $outputPropertyName
                            }
                        }
                )

                $Script:AutotaskReferenceMetaCache[
                    $metadataCacheKey
                ] = @($referenceFields)

                Write-Information "INFO: Wrote $($referenceFields.Count) Reference Field names to local index for '$Resource'."
            }
            catch {
                Write-Warning "Failed to query Reference Field metadata for resource '$Resource' from $fieldsURI : $($_.Exception.Message)"

                # Return the original objects unchanged
                return $returnedItems
            }
        }

        # Exclusions are applied again here because metadata may have been loaded into cache by an earlier invocation without exclusions
        $referenceFields = @(
            $referenceFields |
                Where-Object {
                    $_.FieldName -notin $ExcludeField
                }
        )

        if ($referenceFields.Count -eq 0) {
            return $returnedItems
        }

        # Holds unique reference IDs grouped by target entity
        #
        # Example:
        #
        #   Contracts = HashSet(123, 456)
        #   Resources = HashSet(789, 1011)
        #
        $referenceIDsByEntity = @{}

        foreach ($referenceField in $referenceFields) {
            foreach ($item in $returnedItems) {
                $property = $item.PSObject.Properties[
                    $referenceField.FieldName
                ]

                if ($null -eq $property) {
                    continue
                }

                if (
                    $null -eq $property.Value -or
                    [string]::IsNullOrWhiteSpace(
                        [string]$property.Value
                    )
                ) {
                    continue
                }

                $referenceID = 0L

                if (
                    -not [long]::TryParse(
                        [string]$property.Value,
                        [ref]$referenceID
                    )
                ) {
                    continue
                }

                # Autotask frequently returns zero for an unassigned reference field
                if ($referenceID -le 0) {
                    continue
                }

                $referenceEntityType = [string](
                    $referenceField.ReferenceEntityType
                )

                if (
                    -not $referenceIDsByEntity.ContainsKey(
                        $referenceEntityType
                    )
                ) {
                    $referenceIDsByEntity[
                        $referenceEntityType
                    ] = New-Object `
                        'System.Collections.Generic.HashSet[long]'
                }

                [void]$referenceIDsByEntity[
                    $referenceEntityType
                ].Add($referenceID)
            }
        }

        # Query all referenced records that are not already cached
        foreach (
            $referenceEntityType in $referenceIDsByEntity.Keys
            ) {
                $referenceResource = & $getReferenceResource `
                $referenceEntityType
                
                # Keep cache keys based on the metadata entity type
                $entityCacheKey = $referenceEntityType.Trim().ToLowerInvariant()

            $missingIDs = @(
                foreach (
                    $referenceID in $referenceIDsByEntity[
                        $referenceEntityType
                    ]
                ) {
                    $labelCacheKey = '{0}|{1}|{2}' -f (
                        $baseURIKey
                    ), (
                        $entityCacheKey
                    ), (
                        $referenceID
                    )

                    if (
                        -not $Script:AutotaskReferenceLabelCache.ContainsKey(
                            $labelCacheKey
                        )
                    ) {
                        [long]$referenceID
                    }
                }
            )

            if ($missingIDs.Count -eq 0) {
                Write-Information "INFO: Cached Reference Value index hit for entity '$referenceEntityType'."
                continue
            }

            Write-Information "INFO: Resolving $($missingIDs.Count) reference value(s) from '$referenceResource'."

            for (
                $offset = 0
                $offset -lt $missingIDs.Count
                $offset += $BatchSize
            ) {
                $lastIndex = [Math]::Min(
                    $offset + $BatchSize - 1,
                    $missingIDs.Count - 1
                )

                $idBatch = @(
                    $missingIDs[$offset..$lastIndex]
                )

                $queryURI = '{0}/V1.0/{1}/query' -f (
                    $Script:AutotaskBaseURI.TrimEnd('/')
                    ), (
                        $referenceResource
                        )

                $queryBody = @{
                    MaxRecords = $idBatch.Count
                    filter     = @(
                        @{
                            op    = 'in'
                            field = 'id'
                            value = @($idBatch)
                        }
                    )
                } | ConvertTo-Json -Depth 6

                try {
                    $queryResponse = Invoke-RestMethod `
                        -Method Post `
                        -Uri $queryURI `
                        -Headers $Script:AutotaskAuthHeader `
                        -ContentType 'application/json' `
                        -Body $queryBody

                    $resolvedObjects = @(
                        $queryResponse.items
                    )

                    $returnedIDs = @{}

                    foreach ($resolvedObject in $resolvedObjects) {
                        $resolvedID = & $getPropertyValue `
                            $resolvedObject `
                            'id'

                        if ($null -eq $resolvedID) {
                            continue
                        }

                        $resolvedLabel = & $getReferenceLabel `
                        $referenceResource `
                        $resolvedObject

                        $labelCacheKey = '{0}|{1}|{2}' -f (
                            $baseURIKey
                        ), (
                            $entityCacheKey
                        ), (
                            $resolvedID
                        )

                        $Script:AutotaskReferenceLabelCache[
                            $labelCacheKey
                        ] = $resolvedLabel

                        $returnedIDs["$resolvedID"] = $true
                    }

                    # Cache records that could not be returned. This prevents deleted, inaccessible or invalid IDs from being queried repeatedly throughout the same PowerShell session
                    foreach ($requestedID in $idBatch) {
                        if (
                            -not $returnedIDs.ContainsKey(
                                "$requestedID"
                            )
                        ) {
                            $labelCacheKey = '{0}|{1}|{2}' -f (
                                $baseURIKey
                            ), (
                                $entityCacheKey
                            ), (
                                $requestedID
                            )

                            $Script:AutotaskReferenceLabelCache[
                                $labelCacheKey
                            ] = $null
                        }
                    }
                }
                catch {
                    Write-Warning "Failed to resolve reference entity '$referenceEntityType' using resource '$referenceResource' from $queryURI : $($_.Exception.Message)"
                }
            }
        }

        # Add the resolved values to the original returned objects
        foreach ($item in $returnedItems) {
            foreach ($referenceField in $referenceFields) {
                $idProperty = $item.PSObject.Properties[
                    $referenceField.FieldName
                ]

                if ($null -eq $idProperty) {
                    continue
                }

                $referenceID = 0L

                if (
                    -not [long]::TryParse(
                        [string]$idProperty.Value,
                        [ref]$referenceID
                    ) -or
                    $referenceID -le 0
                ) {
                    continue
                }

                $referenceEntityType = [string](
                    $referenceField.ReferenceEntityType
                )

                $entityCacheKey = $referenceEntityType.Trim(
                ).ToLowerInvariant()

                $labelCacheKey = '{0}|{1}|{2}' -f (
                    $baseURIKey
                ), (
                    $entityCacheKey
                ), (
                    $referenceID
                )

                if (
                    -not $Script:AutotaskReferenceLabelCache.ContainsKey(
                        $labelCacheKey
                    )
                ) {
                    continue
                }

                $resolvedLabel = $Script:AutotaskReferenceLabelCache[
                    $labelCacheKey
                ]

                # Do not add a property when no useful display label could be determined
                if (
                    [string]::IsNullOrWhiteSpace(
                        [string]$resolvedLabel
                    )
                ) {
                    continue
                }

                $outputPropertyName = [string](
                    $referenceField.OutputPropertyName
                )

                $existingProperty = $item.PSObject.Properties[
                    $outputPropertyName
                ]

                if ($null -ne $existingProperty) {
                    if (
                        $OverwriteExisting.IsPresent -or
                        $null -eq $existingProperty.Value -or
                        [string]::IsNullOrWhiteSpace(
                            [string]$existingProperty.Value
                        )
                    ) {
                        $existingProperty.Value = $resolvedLabel
                    }

                    continue
                }

                $item | Add-Member -NotePropertyName $outputPropertyName -NotePropertyValue $resolvedLabel
            }
        }

        return $returnedItems
    }
}