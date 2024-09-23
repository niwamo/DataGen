class FrequencyDict {
    [Collections.Generic.List[string]]$words
    [Collections.Generic.List[int]]$indices
    [int]$Count
    FrequencyDict () {}
    FrequencyDict([Collections.Generic.Dictionary[string, int]]$words) {
        $this.words = $words.Keys
        $this.indices = [Collections.Generic.List[int]]::new()
        $idx = 0
        foreach ($count in $words.Values) {
            foreach ($_ in 1..$count) {
                $this.indices.Add($idx)
            }
            $idx += 1
        }
        $this.Count = $this.indices.Count
    }
    [string] Get([int]$idx) {
        return $this.words[$this.indices[$idx]]
    }
    [string] ToString() {
        return $this | ConvertTo-Json -Depth 2
    }
    static [FrequencyDict] FromString([string]$data) {
        $obj = $data | ConvertFrom-Json
        $out = [FrequencyDict]::new()
        $out.words = $obj.words
        $out.indices = $obj.indices
        $out.Count = $obj.Count
        return $out
    }
}

function New-FDFromString {
    param([string]$Data) return [FrequencyDict]::FromString($Data)
}