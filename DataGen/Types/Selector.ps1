class Getter {
    [Object] Get ([int]$idx) {
        return [Object]
    }
}

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
        $n = $Collection.Count - 1
        return $Collection.Get(($this.generator.Next(0, $n)))
    }
    # $Collection can be anything that implements .Count and .Get()
    [string[]] select ($Collection, [int]$Count) {
        $n = $Collection.Count - 1
        $out = for ($i=0; $i -lt $Count; $i++) { 
            $Collection.Get(($this.generator.Next(0, $n)))
        }
        return $out
    }
    [int] num ([int]$Minimum, [int]$Maximum) {
        return $this.generator.Next($Minimum, $Maximum)
    }
    [int] num ([int]$Maximum) {
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