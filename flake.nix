{
  description = "officeproof: edit Word and Excel documents as an agent and prove the result";

  inputs.nixpkgs.url = "github:NixOS/nixpkgs/nixos-unstable";

  outputs =
    { self, nixpkgs }:
    let
      systems = [
        "aarch64-darwin"
        "x86_64-darwin"
        "aarch64-linux"
        "x86_64-linux"
      ];
      forAll = f: nixpkgs.lib.genAttrs systems (system: f nixpkgs.legacyPackages.${system});
    in
    {
      packages = forAll (pkgs: rec {
        officeproof = pkgs.callPackage ./nix/package.nix { };
        default = officeproof;
      });

      overlays.default = final: _prev: {
        officeproof = final.callPackage ./nix/package.nix { };
      };
    };
}
