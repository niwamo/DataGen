class Selector {
    [System.Random]$generator
    Selector () {
        $this.generator = New-Object System.Random
    }
    Selector ([int]$Seed) {
        $this.generator = New-Object System.Random($Seed)
    }
    # $Collection can be anything that implements .Count and .Get()
    [string] select ($Collection) {
        $n = [Math]::Max(0, $Collection.Count)
        return $Collection.Get(($this.generator.Next(0, $n)))
    }
    # $Collection can be anything that implements .Count and .Get()
    [string[]] select ($Collection, [int]$Count) {
        $n = [Math]::Max(0, $Collection.Count)
        $out = for ($i=0; $i -lt $Count; $i++) { 
            $Collection.Get(($this.generator.Next(0, $n)))
        }
        return $out
    }
    [int] num ([int]$Minimum, [int]$Maximum) {
        if ($Minimum -gt $Maximum) {
            Write-Error "Given min ($Minimum) was greater than the given max ($Maximum)"
            return -1
        }
        return $this.generator.Next($Minimum, $Maximum)
    }
    [int] num ([int]$Maximum) {
        if ($Maximum -lt 0) {
            Write-Error "Given maximum ($Maximum) was less than zero"
            return -1
        }
        return $this.generator.Next(0, $Maximum)
    }
}

function New-Selector {
    param([int]$Seed=$null) 
    if ($Seed) { 
        return [Selector]::new($Seed) 
    } else { 
        return [Selector]::new() 
    }
}