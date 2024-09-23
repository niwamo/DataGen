class SText {
    [Collections.Generic.List[
        ValueTuple[string, Collections.Generic.List[string]]
    ]]$data
    SText () {}
    SText ([PSCustomObject]$InputObject) {
        $this.data = [Collections.Generic.List[
            ValueTuple[string, [Collections.Generic.List[string]]]
        ]]::new()
        $props = ($InputObject | Get-Member -MemberType NoteProperty).Name
        foreach ($prop in $props) {
            $null = $this.data.Add(
                [ValueTuple[string, Collections.Generic.List[string]]]::new(
                    $prop, $InputObject.$prop
                )
            )
        }
    }
    [string[]] GetText([Random]$generator) {
        $out = foreach ($i in 1..$generator.next(1,100)) {
            $i = $generator.next(0, $this.data.Count - 1)
            $j = $generator.next(0, $this.data[$i][1].Count - 1)
            [string]::format(
                "{0}: {1}; ",
                $this.data[$i][0],
                $this.data[$i][1][$j]
            )
        }
        return $out
    }
    [string] ToString() {
        return ($this.data | ConvertTo-Json -Depth 2)
    }
    static [SText] FromString([string]$data) {
        $items = $data | ConvertFrom-Json
        $out = [SText]::new()
        $out.data = [Collections.Generic.List[
            ValueTuple[string, [Collections.Generic.List[string]]]
        ]]::new()
        foreach ($item in $items) {
            $out.data.Add(
                [ValueTuple[string, Collections.Generic.List[string]]]::new(
                    $item.Item1, $item.Item2
                )
            )
        }
        return $out
    }
}

function New-STFromString {
    param([string]$Data) return [SText]::FromString($Data)
}