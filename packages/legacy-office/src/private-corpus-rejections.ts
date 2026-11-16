// Burn-down list of #2263: the private-corpus inputs the legacy readers still
// reject, each with the native rejection it stops at. `private-corpus.test.ts`
// accepts a rejection only when it is listed here with that reason, and fails
// when a listed input renders, so an entry is removed by the change that makes
// its input render (or its reason is updated when the input then stops at a
// later construct).
//
// Inputs are identified by the SHA-256 of their bytes; the corpus is local and
// its file names are not published. Run the corpus test with
// OOXML_LEGACY_CORPUS_DETAILS=1 to print the local paths.
//
// The list ends with only the classes the readers reject by policy: `corrupt`,
// `limit`, `encryption` and `version` (pre-97 formats). `pending` marks an
// input that must render and does not yet.

export type LegacyCorpusRejectionClass = 'pending' | 'corrupt' | 'limit' | 'encryption' | 'version';

export interface LegacyCorpusRejection {
  readonly format: 'doc' | 'xls' | 'ppt';
  readonly class: LegacyCorpusRejectionClass;
  /** The first native rejection of the load, marker included. */
  readonly reason: string;
  /** Lowercase hex SHA-256 of each input that stops at `reason`. */
  readonly sha256: readonly string[];
}

export const LEGACY_CORPUS_REJECTIONS: readonly LegacyCorpusRejection[] = [
  {
    format: 'doc',
    class: 'pending',
    reason: 'UNSUPPORTED:Word drawing group uses an unsupported wrap contour',
    sha256: [
      '0f36731ea46006066142d66e8334f48131e3c401b76206b6136d950379c3d895',
      '3bfa4d91953fbab547bc76e5663871d37b746ddae50b1b1709897aa7247ab834',
      '42f0440366e168e648c8626698b169ebbef7c85d58273d265056d9992677c2bb',
      '44069b21ace87b8dd19ccda3d19cb2d88439f8effd020855f081dde0189da578',
      '7b973095336d57d88615743d1fea6ae3b07842c1b204665d167e7e31f727f97c',
      'cd733747b2c7722f0253c309c1a53eb6c1520e51d0ff15afeeb117aa11e7e937',
      'd9097d7e6f489782b2193cf8148bbfdf4b725d71ec039c8ddeb3d9291a2505c7',
      'f54088c5249a2f341beab79e91624e34602432a407bb2be9a32feb8576b7252b',
    ],
  },
  {
    format: 'doc',
    class: 'pending',
    reason: 'UNSUPPORTED:direct DOC model cannot place a right-to-left table with a preferred indent',
    sha256: [
      '1227926ac4ad4b0dd6698cecf6f0413d3a45af3f5c0bb3c5ba73a8b2af01e2ee',
      '18377936544b941bf36499777e411c92ab004e20f89cf9f5ecf3a7e2af46fddb',
      '3f608f291e44d5009a4ced1a7c13933776e2235b8da0220f9cdc498cfdf046b9',
      '4754ff5d035135b99024c4ab89ff07f2ba08e75c9e015bc0193ba0f716f36011',
      '49b3f77c6676fc80747a9246ed1383ddff5ce7baf76b823cafcfdfa650ae0abe',
      '4a43bd6d3624bc9578793b7d22d2cbf89576198290f6988d99b8a2caf3d4fa3a',
      '4cb65cc553cd72eaa0a82d962698df706f87cefdc207dd12fd687229e7c1f126',
      '56ca11801913e13ac3b4537d61a789dd4d2de265ce828b618f528be2c4bc44a3',
      '83f32ff0a125acc388544b59424e469dc0ac454c16cd0dc905cc9c19427da341',
      '96f1c148496999ab23a230010e9003bb7371f24abdf474af84bb99894b26040e',
      'bcf0110edbc64d99878cf68ed2481b15486509cded8f755e8c6748fd9384705d',
      'cafd9f5e324dfb46656c04b309d5da7ed644a903a2c9861dc951adf7003c1ce6',
      'e2f1db8439b11b8c49a14c07ee346e5087a5062ed524d8d5fcb3ac8fb2d97446',
      'f22173c6c3793778cb147b4fbbdc0cfc8af970c6483b2b08d96072f93b750884',
      'f8933961253ce5cf97f31b817e638af4e9dd587d29d947a955c534f401bbd03c',
    ],
  },
  {
    format: 'doc',
    class: 'pending',
    reason: 'UNSUPPORTED:direct DOC model encountered omitted drawing content',
    sha256: [
      '0e4d16af4fa24a58bdf8ed078bbdd56584c1cf6cc7a4f94225c26eb4800be8ec',
    ],
  },
  {
    format: 'doc',
    class: 'pending',
    reason: 'UNSUPPORTED:direct DOC model encountered unsupported formatting',
    sha256: [
      '0b06056384dbe2cdb3241d4edfe5c14902239b38d7e27fac2bb9dbd54f3fb95e',
      '0ce7bd0ecbb964176d74d9194df85c79edec505b37cc886ee36d5ccbfd0c348d',
      '2088b544152cff03aa118869973342ee18a53cf15aa30ce58bdf8a3000ee8218',
      '252c9f7f747306e33e97bb27d6d9948c75170911e4cf24811611add613d46912',
      '27c0e64eefc3d7c273812d7f5937135b5fe79ee4d5e29d7ad39d76a118fa6679',
      '426bc1f9d80a7b6f5cdb6e2a60c7d7f0c51737127674968da6a5bf25dff15d7d',
      '44f3a4c4cc27ac8041a63d14c3c356200643056d9a92d891e2dcdaa5c223cb80',
      '495e72e7133bd45ca0b282cc5d2887cbda8e8a81bb43822efe204e0c9180386c',
      '50e28ddcd71b4bcdfe554610a6866534a4df659aac6dffe8e54574ae4c3c3d14',
      '59b184611ffcf70f8ab1fb5463b1e676734c3a38ccb547868c09e0b9b3602fc3',
      '64c6b8280ad2dda3e4f6137e2c141d09f2868aefceaa57071fa0c70af2b730df',
      '6bf15ac5d3d9fb8a2670c51c2c5612088e5de8925c4a46ca041346d97a1eee77',
      '77cc0967685829a34d84507e485260e04013fc87c9412ffe4c385c97523b0dc6',
      '78017c9c675b2c3259893eb9b5d8ff3113420578896028c9c596d21f0fe60086',
      '9253eb727b5768a82c6df04afd12d87edca1b2bb4e1cc7345204458a1745fea0',
      '9f851526d3ab28f836db221c94d94e03d49b69d15c84890b6a3d0cacade1ce05',
      'a2671763480bcb25fb537553ed3365fd8297d240637c23fd106e31124c1c4fd6',
      'aa086942f9296b877030f43e721778f28216456b9bf6f845eb6135e7f2d8a469',
      'b8b96a3f65ce64cc53d0e50c7abd228a3e6dfee91c55bf80816bdf4e81fbf82c',
      'ff8316e40ac525d87c6782e6b74248d4194e68777f0f638c58d6012c54bd257c',
    ],
  },
  {
    format: 'doc',
    class: 'pending',
    reason: 'UNSUPPORTED:invalid Word ignored border width',
    sha256: [
      '072a33336e07c06a2b52c7e42aaef641f05ed6e568edb26bcff5e55621a9c5fc',
      '38607ef202c0a44f4ad6e8367bfe8e8a42862bf70ebe61a57acbc52918dd9891',
      '7df84e60327778a241cdddd1db824d7f58b8d278736ede8e20a5a9afe0784d60',
      'ad8ff737de33169fb5183dfad51d4c1d1b867597370fc9aafb92cc2ef75e59d8',
    ],
  },
  {
    format: 'doc',
    class: 'pending',
    reason: 'UNSUPPORTED:invalid Word word-breaking method',
    sha256: [
      '55cdf0f27654a5cf2df8b5e41b36bc83fe1c0469af0ed3bea11be4b982d3a6ba',
      '63c7d7e4d0799b19df6022c419b5b47ac0cdc0e92b12b491fc5d07df3d281a35',
      '826536a36ffdecd3a33e676588530a4b824654b219a809788c02d56e80aae160',
      '8f25f957b2ae11d86b7844fd86f5334212a96636456dec5e77ba881cc030b75e',
      '923d9905004440c7a269a27ca55871f478786f32433a63db39beb74a9193ada2',
    ],
  },
  {
    format: 'doc',
    class: 'pending',
    reason: 'UNSUPPORTED:unsupported Word FIB version',
    sha256: [
      '77ed5f442276fead87ce2acf56296fea03c13e6f8cf8fe7005d4a657ba552b05',
    ],
  },
  {
    format: 'doc',
    class: 'pending',
    reason: 'UNSUPPORTED:unsupported Word border type 0xFF ignore semantics',
    sha256: [
      '0d11a8c12dc7fa5467d05afdfc18fde0d3e331003aff12958112b2a171cbfb18',
      '18282c1a498e653022bf8328658bf1653a0541d0743cbe1d9bd3302fa741029c',
      '424db181f0cd84aecc9b99ecb2a7141f015504af37916406cb42b63a85aaf87f',
      '4d8791e6902f5cd000501de13c2af44b75d262a7d9542bcf821362061bd9ce4f',
      '64387e224f88c1fc997d4cea9189e7b5e19a983d1fc9ac1383408a841750a016',
      'ca61453392178193baba7969ea91d273b88dcaf3c0a639e663adb80513cdf905',
      'e70ca483521576bb3507a30262808f2d50e8b9bf21bbcd54f9c2749cdcf7fcc4',
    ],
  },
  {
    format: 'ppt',
    class: 'pending',
    reason: 'UNSUPPORTED:PowerPoint alternative shape XML custom geometry cannot be compared with the binary shape',
    sha256: [
      'da839619e16c35e281c6e66831c5fe0d12640608ddbfd7137bcad1d2afde4206',
    ],
  },
  {
    format: 'ppt',
    class: 'pending',
    reason: 'UNSUPPORTED:PowerPoint alternative shape XML fill cannot be compared with the binary shape',
    sha256: [
      '1180be153c16c3e444edc565b4a5a8ea810e9d805035eb91243eee0d5f1a4902',
      '2bdd29055fbc977180c8631409d3489e6536af01002a28954b58f0f6f3044ae5',
      '2dd5f2477c2018c42459f98593628b1465d0574c31a46c2ba675426a2ce96366',
      '4029e8998391b1942bea93eaa77ab602f66681e2e76b2b55f8caa6b7da68150a',
      '5c4cb512fc80825032ae22b0f999134b3ed378bb0ea6d3bed434e8a7307a66b2',
      '735e88e51c18c8bc6c0fea612cbbe990cf65667b7dc1d9dac07fda0d0317c86f',
      'a8175cc42a961767bb38d30bc610a8c75237c6eafe28fe803b6719fa68321b2c',
    ],
  },
];
