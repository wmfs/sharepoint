# [1.18.0](https://github.com/wmfs/sharepoint/compare/v1.17.0...v1.18.0) (2023-03-13)


### 🛠 Builds

* **deps-dev:** update dependency semantic-release to v20.1.1 ([884755a](https://github.com/wmfs/sharepoint/commit/884755a3badc6e51915b5cf0b5d625b75d1b4ff5))
* **deps:** update dependency axios to v1.3.4 ([6059e28](https://github.com/wmfs/sharepoint/commit/6059e282480949bd9d94c7c404d47d03db7e374d))

# [1.17.0](https://github.com/wmfs/sharepoint/compare/v1.16.0...v1.17.0) (2023-02-17)


### 🛠 Builds

* **deps-dev:** update dependency [@semantic-release](https://github.com/semantic-release)/changelog to v6.0.2 ([47b9330](https://github.com/wmfs/sharepoint/commit/47b9330dbc79921ab39a760cdf39ed6910b9d0a1))
* **deps-dev:** update dependency chai to v4.3.7 ([b2e9805](https://github.com/wmfs/sharepoint/commit/b2e98058853f55f0a891aa47b8dde3cc3808244d))
* **deps-dev:** update dependency dotenv to v16.0.3 ([bbbc34c](https://github.com/wmfs/sharepoint/commit/bbbc34caaa0b2b3168c82766880ac038ffb0337f))
* **deps-dev:** update dependency mocha to v10.2.0 ([cd99169](https://github.com/wmfs/sharepoint/commit/cd99169295d121516969576a91d7412db3ab4560))
* **deps-dev:** update dependency semantic-release to v20 ([7a5467c](https://github.com/wmfs/sharepoint/commit/7a5467c61752aeb9f350a2705e51a89dfcc8e05d))
* **deps-dev:** update dependency semantic-release to v20.0.1 ([6b84351](https://github.com/wmfs/sharepoint/commit/6b8435139a5b0ad565ec1da824143b773af2c892))
* **deps-dev:** update dependency semantic-release to v20.0.2 ([32451ae](https://github.com/wmfs/sharepoint/commit/32451aeea7af63db8be8787e630f3ef753b86642))
* **deps-dev:** update dependency semantic-release to v20.0.3 ([d77b3ef](https://github.com/wmfs/sharepoint/commit/d77b3ef8430ed06fb0f78be8939d6671481ed307))
* **deps-dev:** update dependency semantic-release to v20.0.4 ([623d2a3](https://github.com/wmfs/sharepoint/commit/623d2a35c7702cd1d614f75254536dca1ea31edc))
* **deps-dev:** update dependency semantic-release to v20.1.0 ([87f386f](https://github.com/wmfs/sharepoint/commit/87f386fd4832b8bd37cba09701aa2d2df4099e06))
* **deps:** update dependency axios to v1.3.3 ([b65b69d](https://github.com/wmfs/sharepoint/commit/b65b69d617cbcc3a1c5e3ab1d1d55e3a56026309))

# [1.16.0](https://github.com/wmfs/sharepoint/compare/v1.15.0...v1.16.0) (2022-11-02)


### 🛠 Builds

* **deps-dev:** update dependency mocha to v10.1.0 ([8d3eb03](https://github.com/wmfs/sharepoint/commit/8d3eb03f73675850bc28cd5a6bdbb0ed77b620f4))
* **deps-dev:** update dependency semantic-release to v19.0.5 ([156557e](https://github.com/wmfs/sharepoint/commit/156557ec3d75c07e7a33518c612c075adaef7407))
* **deps:** update dependency axios to v1 ([50ddae7](https://github.com/wmfs/sharepoint/commit/50ddae7e425ad1c3b8b26151cd351a3e00456d49))

# [1.15.0](https://github.com/wmfs/sharepoint/compare/v1.14.0...v1.15.0) (2022-11-02)


### 🛠 Builds

* **deps:** update dependency axios to v0.27.2 ([db5e251](https://github.com/wmfs/sharepoint/commit/db5e2514937f451f6e64817898e83581a49a80be))

# [1.14.0](https://github.com/wmfs/sharepoint/compare/v1.13.0...v1.14.0) (2022-11-02)


### 🛠 Builds

* **deps:** update dependency node-sp-auth to v3.0.7 ([e3e0525](https://github.com/wmfs/sharepoint/commit/e3e05258e8eaed019a03d4ec2afeaa9aeb959772))

# [1.13.0](https://github.com/wmfs/sharepoint/compare/v1.12.2...v1.13.0) (2022-11-02)


### 🛠 Builds

* **deps:** update dependency node-sp-auth to v3.0.5 ([57d5c30](https://github.com/wmfs/sharepoint/commit/57d5c30fc00a2a63679f2449fb0e9b6a7d3344d8))

## [1.12.2](https://github.com/wmfs/sharepoint/compare/v1.12.1...v1.12.2) (2022-08-30)


### 🐛 Bug Fixes

* encode paths and filenames passed from options into axios url requests ([af1ff1c](https://github.com/wmfs/sharepoint/commit/af1ff1cc859c9b73f7b5e7325b731455df1d218e))


### 🛠 Builds

* **deps-dev:** update dependency dotenv to v16.0.1 ([2f40b38](https://github.com/wmfs/sharepoint/commit/2f40b38df345e555d986fef988cdea6cf1bc4387))
* **deps-dev:** update dependency mocha to v10 ([0a09f62](https://github.com/wmfs/sharepoint/commit/0a09f6232dbae06b4a81713a24cbb07e76617759))
* **deps-dev:** update dependency semantic-release to v19.0.3 ([b12cdbb](https://github.com/wmfs/sharepoint/commit/b12cdbbdc891030705014425029cd02452d7c3eb))
* **deps-dev:** update dependency standard to v17 ([b181ecc](https://github.com/wmfs/sharepoint/commit/b181ecc76b402d2c3b61005a99d8f4c38376c1b8))

## [1.12.1](https://github.com/wmfs/sharepoint/compare/v1.12.0...v1.12.1) (2022-03-29)


### 🐛 Bug Fixes

* if the whole file was uploaded in the first chunk then commit (finish) the upload, otherwise it will never been resolved/committed. Also return info about the created file ([9310b28](https://github.com/wmfs/sharepoint/commit/9310b28fe97321e95b9b02814c78ec68b9d84998))


### 🛠 Builds

* **deps-dev:** update dependency chai to v4.3.5 ([5e1ae7c](https://github.com/wmfs/sharepoint/commit/5e1ae7c2ce74d39ea6e0186e2c3b67c9a1767b02))
* **deps-dev:** update dependency chai to v4.3.6 ([609a75b](https://github.com/wmfs/sharepoint/commit/609a75be6d0aba130f350fb35b49a1176fa8e319))
* **deps-dev:** update dependency dotenv to v14.3.2 ([e281620](https://github.com/wmfs/sharepoint/commit/e281620bb59d94d74947486a6eba3b6df18dc98a))
* **deps-dev:** update dependency dotenv to v15 ([527e84b](https://github.com/wmfs/sharepoint/commit/527e84b2665836ad22a6635285a12f97587b47ac))
* **deps-dev:** update dependency dotenv to v15.0.1 ([d58b9ab](https://github.com/wmfs/sharepoint/commit/d58b9ab74b7f8cdf88670c0fe7a45a047ff53953))
* **deps-dev:** update dependency dotenv to v16 ([19d8fd9](https://github.com/wmfs/sharepoint/commit/19d8fd95c98cf8f5bb1dbab596bfb864d88ffdae))
* **deps-dev:** update dependency mocha to v9.2.1 ([566a139](https://github.com/wmfs/sharepoint/commit/566a139f3e4a5159a741a30852e531edd1c39462))
* **deps-dev:** update dependency mocha to v9.2.2 ([40c1372](https://github.com/wmfs/sharepoint/commit/40c13729387b0b8ef64adfeaf7e851ae6d0e9f19))

# [1.12.0](https://github.com/wmfs/sharepoint/compare/v1.11.0...v1.12.0) (2022-01-25)


### 🛠 Builds

* **deps:** update dependency axios to v0.25.0 ([ec95a77](https://github.com/wmfs/sharepoint/commit/ec95a773517a95d68aa4be04d2cf9947118d07de))

# [1.11.0](https://github.com/wmfs/sharepoint/compare/v1.10.0...v1.11.0) (2022-01-25)


### 🛠 Builds

* **deps-dev:** update dependency [@semantic-release](https://github.com/semantic-release)/changelog to v6 ([8eae8de](https://github.com/wmfs/sharepoint/commit/8eae8de7992c05bb74f39a5cf344fce323f94d38))
* **deps-dev:** update dependency [@semantic-release](https://github.com/semantic-release)/git to v10 ([98835e5](https://github.com/wmfs/sharepoint/commit/98835e5751bdfd6436b4ecfe7b8d9f6de4ee5eb5))
* **deps-dev:** update dependency dotenv to v11 ([2a53c9a](https://github.com/wmfs/sharepoint/commit/2a53c9ad410a69c56dc2c3c54d0984ae78b212e0))
* **deps-dev:** update dependency dotenv to v12 ([f659ad8](https://github.com/wmfs/sharepoint/commit/f659ad834bbc079ca093f0c4bd019983cd61b1a1))
* **deps-dev:** update dependency dotenv to v12.0.4 ([f561cba](https://github.com/wmfs/sharepoint/commit/f561cba770f7c41454ea706edbed6664ec4ddc24))
* **deps-dev:** update dependency dotenv to v13 ([299f403](https://github.com/wmfs/sharepoint/commit/299f4037b08ac644642004eec5098421ae5bafaf))
* **deps-dev:** update dependency dotenv to v14 ([1c6896e](https://github.com/wmfs/sharepoint/commit/1c6896e54b41b21726161d83df225c7cb60fac55))
* **deps-dev:** update dependency dotenv to v14.2.0 ([96603a3](https://github.com/wmfs/sharepoint/commit/96603a3a27fbbf2e3616f21ee560ccade22ae236))
* **deps-dev:** update dependency dotenv to v14.3.0 ([e409d41](https://github.com/wmfs/sharepoint/commit/e409d419472f19e88e7d694831b803a8fc7c3e16))
* **deps-dev:** update dependency mocha to v9.1.2 ([cfb48e4](https://github.com/wmfs/sharepoint/commit/cfb48e40279e4e6dc36d2f31b7ba182bdb5edf9f))
* **deps-dev:** update dependency mocha to v9.1.3 ([ec3b431](https://github.com/wmfs/sharepoint/commit/ec3b4317797531fafe4d601a84fbbeee5a7693a3))
* **deps-dev:** update dependency mocha to v9.1.4 ([7afcdac](https://github.com/wmfs/sharepoint/commit/7afcdacb6b59731ceeef5225351fe4049c2c2740))
* **deps-dev:** update dependency mocha to v9.2.0 ([d1fe594](https://github.com/wmfs/sharepoint/commit/d1fe5947735915efe2bf9c596ca10e6ae4800a70))
* **deps-dev:** update dependency semantic-release to v18 ([8fac748](https://github.com/wmfs/sharepoint/commit/8fac748d8f677bb31105b78cd5bcc9d23bba63b6))
* **deps-dev:** update dependency semantic-release to v18.0.1 ([8062b84](https://github.com/wmfs/sharepoint/commit/8062b84907a3576756a175f43b7485665b3ce2f7))
* **deps-dev:** update dependency semantic-release to v19 ([09d3d24](https://github.com/wmfs/sharepoint/commit/09d3d242e52393c68894db2732abfd75f630e236))
* **deps-dev:** update dependency standard to v16.0.4 ([45f00b5](https://github.com/wmfs/sharepoint/commit/45f00b516913974fb54ec9b86ad6b6427bdc7f3e))
* **deps-dev:** update semantic-release monorepo ([0348a83](https://github.com/wmfs/sharepoint/commit/0348a83bd958ff15f99a58fa61b89a61043a3faa))
* **deps:** update dependency node-sp-auth to v3.0.4 ([319b53d](https://github.com/wmfs/sharepoint/commit/319b53d7e2b7a99688e4f1075b50e9de0be1b4a0))

# [1.10.0](https://github.com/wmfs/sharepoint/compare/v1.9.0...v1.10.0) (2021-09-10)


### 🛠 Builds

* **deps:** update dependency axios to v0.21.2 [security] ([7ac9cb4](https://github.com/wmfs/sharepoint/commit/7ac9cb42ca8e8457cb560fb1245535254cff7be6))
* **deps-dev:** bump chai from 4.3.0 to 4.3.1 ([5bf3ebe](https://github.com/wmfs/sharepoint/commit/5bf3ebe965d8de9820709a074a105aeb7bc4f6d7))
* **deps-dev:** bump chai from 4.3.1 to 4.3.2 ([d814c15](https://github.com/wmfs/sharepoint/commit/d814c158df4d435f1167e988a0de7164e85cc0cf))
* **deps-dev:** bump chai from 4.3.2 to 4.3.3 ([c0ed7af](https://github.com/wmfs/sharepoint/commit/c0ed7af3d24acb3d87177cfd5b03cfaf2c32613f))
* **deps-dev:** bump chai from 4.3.3 to 4.3.4 ([aed3966](https://github.com/wmfs/sharepoint/commit/aed3966f2b68467aec6db63a5fd140fe0a187476))
* **deps-dev:** bump codecov from 3.8.1 to 3.8.2 ([018d824](https://github.com/wmfs/sharepoint/commit/018d8249b08a23b176ffaa351f224e30a9402a14))
* **deps-dev:** bump codecov from 3.8.2 to 3.8.3 ([7053497](https://github.com/wmfs/sharepoint/commit/705349740e90b1806752b781619797a32801a4cf))
* **deps-dev:** bump dotenv from 8.6.0 to 9.0.0 ([a905156](https://github.com/wmfs/sharepoint/commit/a905156c4dfa59cc7b313427915f9ea44c97a109))
* **deps-dev:** bump dotenv from 9.0.2 to 10.0.0 ([684e561](https://github.com/wmfs/sharepoint/commit/684e5610ee29f52b4ccad9f222e2030704b4f158))
* **deps-dev:** bump mocha from 8.3.0 to 8.3.1 ([ff3ac2b](https://github.com/wmfs/sharepoint/commit/ff3ac2b5fdce64250dc87205e1b6ca9e375f138f))
* **deps-dev:** bump mocha from 8.3.1 to 8.3.2 ([1b99b49](https://github.com/wmfs/sharepoint/commit/1b99b4966cf551b8499d800417e76230a631af15))
* **deps-dev:** bump mocha from 8.3.2 to 8.4.0 ([03a50d8](https://github.com/wmfs/sharepoint/commit/03a50d8eef45a19efaf1fe1ba335c8421a050011))
* **deps-dev:** bump mocha from 8.4.0 to 9.0.0 ([bdac9a6](https://github.com/wmfs/sharepoint/commit/bdac9a692ddbc2951056d7a60ab1601969e3a8c6))
* **deps-dev:** bump mocha from 9.0.0 to 9.0.1 ([b56e9b5](https://github.com/wmfs/sharepoint/commit/b56e9b5eb62440118b3ceb10539285d609759f23))
* **deps-dev:** bump mocha from 9.0.1 to 9.0.2 ([e308144](https://github.com/wmfs/sharepoint/commit/e30814454214b749f2a037ce92cdc3baa7b5d689))
* **deps-dev:** bump mocha from 9.0.2 to 9.0.3 ([f29711d](https://github.com/wmfs/sharepoint/commit/f29711d4d7a451132ca9a0abb559c19790fefd39))
* **deps-dev:** bump semantic-release from 17.3.9 to 17.4.0 ([905922b](https://github.com/wmfs/sharepoint/commit/905922bf72c1670e2de83871b2557dac2b32e90f))
* **deps-dev:** bump semantic-release from 17.4.0 to 17.4.1 ([0e328b4](https://github.com/wmfs/sharepoint/commit/0e328b4164aeb51c9a4391f7a84666f85191bc7e))
* **deps-dev:** bump semantic-release from 17.4.1 to 17.4.2 ([b25e508](https://github.com/wmfs/sharepoint/commit/b25e508b1abcf1c5cbce4709daadc146370288a9))
* **deps-dev:** bump semantic-release from 17.4.2 to 17.4.3 ([61fb5e3](https://github.com/wmfs/sharepoint/commit/61fb5e3ed1b338dadfc03cb6a2a1715dd8a14a62))
* **deps-dev:** bump semantic-release from 17.4.3 to 17.4.4 ([163a2d7](https://github.com/wmfs/sharepoint/commit/163a2d7573b2ba8ae32a3bf47bb4ce5ec9f2ddc0))
* **deps-dev:** pin dependency dotenv to 10.0.0 ([c65dfd9](https://github.com/wmfs/sharepoint/commit/c65dfd9422e0b97b787badbfdb5d207dcb7c49d5))
* **deps-dev:** update dependency [@semantic-release](https://github.com/semantic-release)/git to v9.0.1 ([3c746fb](https://github.com/wmfs/sharepoint/commit/3c746fb73930df8fbee906061958eeac79aa2526))
* **deps-dev:** update dependency mocha to v9.1.0 ([8cb533e](https://github.com/wmfs/sharepoint/commit/8cb533ea36658fa732957ce7204de91fdee5f38e))
* **deps-dev:** update dependency mocha to v9.1.1 ([718fc8f](https://github.com/wmfs/sharepoint/commit/718fc8ff121ad9bfbb1f2a15de202660f4119fc0))
* **deps-dev:** update dependency semantic-release to v17.4.5 ([c3aa2dc](https://github.com/wmfs/sharepoint/commit/c3aa2dc429508b0cd0e5759c8960a6219cf22770))
* **deps-dev:** update dependency semantic-release to v17.4.6 ([6421b8a](https://github.com/wmfs/sharepoint/commit/6421b8ae4c4df18434394d97c9749c59c380091f))
* **deps-dev:** update dependency semantic-release to v17.4.7 ([a2ffb24](https://github.com/wmfs/sharepoint/commit/a2ffb2403d3b0d5aa0278743a50d7f837c184d39))


### ♻️ Chores

* add renovate config [ch6600] ([b8bfead](https://github.com/wmfs/sharepoint/commit/b8bfeadee077873128809c7b1c66f2c2f468589e))

# [1.9.0](https://github.com/wmfs/sharepoint/compare/v1.8.0...v1.9.0) (2021-02-24)


### 🛠 Builds

* **deps:** bump node-sp-auth from 3.0.1 to 3.0.3 ([d9bc674](https://github.com/wmfs/sharepoint/commit/d9bc6743aa6b2842857eeaaec6c6698e7fb552dc))
* **deps-dev:** bump chai from 4.2.0 to 4.3.0 ([cd2ea79](https://github.com/wmfs/sharepoint/commit/cd2ea795460bcf28576f974cda93eb6808fa62f0))
* **deps-dev:** bump mocha from 8.2.1 to 8.3.0 ([97c1d44](https://github.com/wmfs/sharepoint/commit/97c1d446751ed9d8eb034b6233acb753f0f52ac8))
* **deps-dev:** bump semantic-release from 17.3.1 to 17.3.2 ([ea61a07](https://github.com/wmfs/sharepoint/commit/ea61a072f82fb3eb47bb88b9842c7ec6e6d95aca))
* **deps-dev:** bump semantic-release from 17.3.2 to 17.3.3 ([f027621](https://github.com/wmfs/sharepoint/commit/f027621911d66abdc4c194658d5a9515f1138638))
* **deps-dev:** bump semantic-release from 17.3.3 to 17.3.4 ([1d2523b](https://github.com/wmfs/sharepoint/commit/1d2523bc2b19316e41557e1e9db0ea9dbdd41775))
* **deps-dev:** bump semantic-release from 17.3.4 to 17.3.5 ([38386c3](https://github.com/wmfs/sharepoint/commit/38386c3833185c6adcb1d420936683079618399a))
* **deps-dev:** bump semantic-release from 17.3.5 to 17.3.6 ([53d453b](https://github.com/wmfs/sharepoint/commit/53d453bc739d55acdd626062041a0c8341d6719a))
* **deps-dev:** bump semantic-release from 17.3.6 to 17.3.7 ([c2677fd](https://github.com/wmfs/sharepoint/commit/c2677fd77fe496b6e6318efdce28a1aa8d87aad9))
* **deps-dev:** bump semantic-release from 17.3.7 to 17.3.8 ([b7edbb6](https://github.com/wmfs/sharepoint/commit/b7edbb62c452eeeaa8f3e151ad265da31e4b29fa))
* **deps-dev:** bump semantic-release from 17.3.8 to 17.3.9 ([29f993a](https://github.com/wmfs/sharepoint/commit/29f993abaf26f7be976b01b747360e77c6ff271a))

# [1.8.0](https://github.com/wmfs/sharepoint/compare/v1.7.0...v1.8.0) (2021-01-05)


### 🛠 Builds

* **deps:** bump axios from 0.21.0 to 0.21.1 ([92620ec](https://github.com/wmfs/sharepoint/commit/92620ec1f49203da6db12a2118dc7b4beb2a61ef))
* **deps-dev:** bump codecov from 3.8.0 to 3.8.1 ([0748d1d](https://github.com/wmfs/sharepoint/commit/0748d1d0f28d11a33b479b95835c59d7160a94e5))
* **deps-dev:** bump mocha from 8.2.0 to 8.2.1 ([c775d70](https://github.com/wmfs/sharepoint/commit/c775d7085ee3bddb16d045dfa2f25d2019d42e46))
* **deps-dev:** bump semantic-release from 17.2.1 to 17.2.2 ([ffdd4d0](https://github.com/wmfs/sharepoint/commit/ffdd4d0cadcb3db33e09ad5f405dee1d52918ed8))
* **deps-dev:** bump semantic-release from 17.2.2 to 17.2.3 ([4794f51](https://github.com/wmfs/sharepoint/commit/4794f51c8efe1db4286cc7b2fb41bd8f5cda65a6))
* **deps-dev:** bump semantic-release from 17.2.3 to 17.2.4 ([d884205](https://github.com/wmfs/sharepoint/commit/d88420595cab4a7cfdbfe262254e6d1dbba56b54))
* **deps-dev:** bump semantic-release from 17.2.4 to 17.3.0 ([913b9e5](https://github.com/wmfs/sharepoint/commit/913b9e5b7146b611fb60d7430d16111f7ae79867))
* **deps-dev:** bump semantic-release from 17.3.0 to 17.3.1 ([d933549](https://github.com/wmfs/sharepoint/commit/d933549309540910616de736a96ea8eebcf9f900))
* **deps-dev:** bump standard from 15.0.0 to 15.0.1 ([e22da97](https://github.com/wmfs/sharepoint/commit/e22da97ed2440bf3e162b48aa27c409f3448e97e))
* **deps-dev:** bump standard from 15.0.1 to 16.0.0 ([bfeb14c](https://github.com/wmfs/sharepoint/commit/bfeb14c6c531fc1b7e5300ca911416cb99413ee6))
* **deps-dev:** bump standard from 16.0.0 to 16.0.1 ([5ed6438](https://github.com/wmfs/sharepoint/commit/5ed6438a303d02dd8bd9621223a7a9772f6bdbe2))
* **deps-dev:** bump standard from 16.0.1 to 16.0.2 ([819f58c](https://github.com/wmfs/sharepoint/commit/819f58cfc1add1e3d43a6d78f965134331adcf3f))
* **deps-dev:** bump standard from 16.0.2 to 16.0.3 ([d511259](https://github.com/wmfs/sharepoint/commit/d51125923d08af914507f92a5b340cd87e191f45))


### ⚙️ Continuous Integrations

* **circle:** separate linting job [ch1009] ([35fd615](https://github.com/wmfs/sharepoint/commit/35fd615b0370475bff8b701efbbc6200c2d98ec7))
* **circle:** update build environment variable context name [ch2771] ([9a82d6b](https://github.com/wmfs/sharepoint/commit/9a82d6b86fc34a5d59a252dd2b608569f66d65c8))

# [1.7.0](https://github.com/wmfs/sharepoint/compare/v1.6.0...v1.7.0) (2020-10-26)


### 🛠 Builds

* **deps:** bump axios from 0.20.0 to 0.21.0 ([c3b71ed](https://github.com/wmfs/sharepoint/commit/c3b71edf79ad295d2aa520dc211f7edf3cfb02a4))
* **deps-dev:** bump codecov from 3.7.2 to 3.8.0 ([282a833](https://github.com/wmfs/sharepoint/commit/282a8334d68980a65075199fbb77d47fc83cd492))
* **deps-dev:** bump cz-conventional-changelog from 3.2.0 to 3.2.1 ([88a2fca](https://github.com/wmfs/sharepoint/commit/88a2fca2b5f9ac867561c3e8a11bd60870a4a482))
* **deps-dev:** bump cz-conventional-changelog from 3.2.1 to 3.3.0 ([3f124a6](https://github.com/wmfs/sharepoint/commit/3f124a64753b29df8e5cd43135b2b087e01d58b3))
* **deps-dev:** bump mocha from 8.1.1 to 8.1.2 ([3282ad9](https://github.com/wmfs/sharepoint/commit/3282ad945fed5c9b29650d4446c80ab8c02c3e22))
* **deps-dev:** bump mocha from 8.1.2 to 8.1.3 ([06e2a95](https://github.com/wmfs/sharepoint/commit/06e2a953df8fa7c57dec66da3f95f20c55e2a193))
* **deps-dev:** bump mocha from 8.1.3 to 8.2.0 ([9580421](https://github.com/wmfs/sharepoint/commit/9580421ac8f6bc9936171850e77b9983d2fe07e8))
* **deps-dev:** bump semantic-release from 17.1.1 to 17.1.2 ([cfe5b69](https://github.com/wmfs/sharepoint/commit/cfe5b69e1669f058a29e76205f9c5f939a4e0f5d))
* **deps-dev:** bump semantic-release from 17.1.2 to 17.2.0 ([e0e24df](https://github.com/wmfs/sharepoint/commit/e0e24df3ef0fd968c41ccd08b4ea6238b814f891))
* **deps-dev:** bump semantic-release from 17.2.0 to 17.2.1 ([a4ce34e](https://github.com/wmfs/sharepoint/commit/a4ce34e3c6bd737f4b952da1d84892d7282f79fc))
* **deps-dev:** bump standard from 14.3.4 to 15.0.0 ([f295c1a](https://github.com/wmfs/sharepoint/commit/f295c1a426c43aa405500e3c00374f8b6136fc10))


### ⚙️ Continuous Integrations

* **circle:** authenticate Docker image pull [ch2767] ([ac1c101](https://github.com/wmfs/sharepoint/commit/ac1c10130729af1e743ab209610bfeea12499c94))
* **circle:** cache dependencies [ch2770] ([ed802d3](https://github.com/wmfs/sharepoint/commit/ed802d3bd7b88b720a408b457e30ec7d25e1f3e7))

# [1.6.0](https://github.com/wmfs/sharepoint/compare/v1.5.0...v1.6.0) (2020-08-24)


### 🛠 Builds

* **deps:** bump axios from 0.19.2 to 0.20.0 ([f4a60a8](https://github.com/wmfs/sharepoint/commit/f4a60a874a7276bb7c0caad6e6e029e80b92995e))
* **deps-dev:** bump codecov from 3.7.0 to 3.7.1 ([e8701c9](https://github.com/wmfs/sharepoint/commit/e8701c91d04462b334132e8d4857464254050c2b))
* **deps-dev:** bump codecov from 3.7.1 to 3.7.2 ([2d885a1](https://github.com/wmfs/sharepoint/commit/2d885a1d3623ae0e064be8ca942d96131aef09a1))
* **deps-dev:** bump mocha from 8.0.1 to 8.1.0 ([c4ccb47](https://github.com/wmfs/sharepoint/commit/c4ccb474dddbd671d970adbfb660b00491ecf41c))
* **deps-dev:** bump mocha from 8.1.0 to 8.1.1 ([5f49204](https://github.com/wmfs/sharepoint/commit/5f492045fdeb6f742ebe63771e7832f44d4027e5))


### ⚙️ Continuous Integrations

* **circle:** separate lint job [ch1009] ([f69f339](https://github.com/wmfs/sharepoint/commit/f69f339a0c16c34f08f234a0b1de0f3f42d0ada4))

# [1.5.0](https://github.com/wmfs/sharepoint/compare/v1.4.1...v1.5.0) (2020-07-16)


### 🛠 Builds

* **deps:** bump node-sp-auth from 2.5.9 to 3.0.1 ([9f9b973](https://github.com/wmfs/sharepoint/commit/9f9b9731865392b36e7ad27d8f5374319408ea96))

## [1.4.1](https://github.com/wmfs/sharepoint/compare/v1.4.0...v1.4.1) (2020-06-30)


### 🐛 Bug Fixes

* No-op README change to try and get semantic-release to run a release. ([03de8af](https://github.com/wmfs/sharepoint/commit/03de8af847c086b053177f25fc369b391ada853b))


### 🛠 Builds

* **deps-dev:** bump mocha from 7.2.0 to 8.0.1 ([ea9b10d](https://github.com/wmfs/sharepoint/commit/ea9b10da4c83b99eeddd4354ae57e8373f20474f))
* **deps-dev:** bump semantic-release from 17.0.8 to 17.1.0 ([8ca1674](https://github.com/wmfs/sharepoint/commit/8ca16742b80f7b8cc9aa36e9e9ff55e5db5f6f81))
* **deps-dev:** bump semantic-release from 17.1.0 to 17.1.1 ([58debd5](https://github.com/wmfs/sharepoint/commit/58debd5850175ef24375b09ce424cfa72792f78f))


### ⚙️ Continuous Integrations

* **circle:** use updated circle node image [skip ci] ([9f36e7f](https://github.com/wmfs/sharepoint/commit/9f36e7f5e2e98e54ad4847877378164d6da2909f))

# [1.4.0](https://github.com/wmfs/sharepoint/compare/v1.3.0...v1.4.0) (2020-06-01)


### 🛠 Builds

* **deps:** bump node-sp-auth from 2.5.7 to 2.5.9 ([c425ace](https://github.com/wmfs/sharepoint/commit/c425acee6fd7e0baddbe725018941c91e310ea20))
* **deps-dev:** bump [@semantic-release](https://github.com/semantic-release)/changelog from 3.0.6 to 5.0.0 ([abfa249](https://github.com/wmfs/sharepoint/commit/abfa249c94d5f642fb42794cb63753490c8c6c1d))
* **deps-dev:** bump [@semantic-release](https://github.com/semantic-release)/changelog from 5.0.0 to 5.0.1 ([16a51cd](https://github.com/wmfs/sharepoint/commit/16a51cdcf412935ede88632b52093f155fc331c3))
* **deps-dev:** bump [@semantic-release](https://github.com/semantic-release)/git from 7.0.18 to 9.0.0 ([9e91abf](https://github.com/wmfs/sharepoint/commit/9e91abf6f5798bdf0fb705e70c2bead49743c25d))
* **deps-dev:** bump codecov from 3.6.1 to 3.6.2 ([09f6563](https://github.com/wmfs/sharepoint/commit/09f65639b5ce788ffb7b0b04abc5e89a79678504))
* **deps-dev:** bump codecov from 3.6.2 to 3.6.3 ([e5d1e72](https://github.com/wmfs/sharepoint/commit/e5d1e725182045abeda00b48d4a2bf247b3c8dcd))
* **deps-dev:** bump codecov from 3.6.3 to 3.6.4 ([d45fbb9](https://github.com/wmfs/sharepoint/commit/d45fbb9ed3dc87c9e859b1fd7df88c0344ce2c4f))
* **deps-dev:** bump codecov from 3.6.4 to 3.6.5 ([ca0d92c](https://github.com/wmfs/sharepoint/commit/ca0d92c064451376d5a290d0d7bf72ec22af750c))
* **deps-dev:** bump codecov from 3.6.5 to 3.7.0 ([4400984](https://github.com/wmfs/sharepoint/commit/4400984a9a58b4a94691bb01322c3895f2023ce9))
* **deps-dev:** bump conventional-changelog-metahub from 4.0.0 to 4.0.1 ([2d53a97](https://github.com/wmfs/sharepoint/commit/2d53a97203f614b33ada6fd3ef98642099eac7c8))
* **deps-dev:** bump cz-conventional-changelog from 3.0.2 to 3.1.0 ([36c7cab](https://github.com/wmfs/sharepoint/commit/36c7cab9011a046d4c555157a3497cfdea516b0b))
* **deps-dev:** bump cz-conventional-changelog from 3.1.0 to 3.2.0 ([384b207](https://github.com/wmfs/sharepoint/commit/384b207edd1bfbdb05c7cd855f4c412f88f51f98))
* **deps-dev:** bump mocha from 7.0.0 to 7.0.1 ([cae62ec](https://github.com/wmfs/sharepoint/commit/cae62ec687be297453fbae7cbe8c4fc15d01dec5))
* **deps-dev:** bump mocha from 7.0.1 to 7.1.0 ([7bc7e97](https://github.com/wmfs/sharepoint/commit/7bc7e97e6d15f7626e5f5a9df08b083626c8216e))
* **deps-dev:** bump mocha from 7.1.0 to 7.1.1 ([e9a5787](https://github.com/wmfs/sharepoint/commit/e9a5787a2d1fd575c89c9a2e8b9268b04ddbe256))
* **deps-dev:** bump mocha from 7.1.1 to 7.1.2 ([5b442f2](https://github.com/wmfs/sharepoint/commit/5b442f203fee98abda5879f94cc197782c240a75))
* **deps-dev:** bump mocha from 7.1.2 to 7.2.0 ([f9b4bd3](https://github.com/wmfs/sharepoint/commit/f9b4bd3fc78e4c09736f10c261405bf1c585ea63))
* **deps-dev:** bump nyc from 15.0.0 to 15.0.1 ([4745085](https://github.com/wmfs/sharepoint/commit/4745085bd88aef2aa1302bd84205dde763426920))
* **deps-dev:** bump nyc from 15.0.1 to 15.1.0 ([76f30b0](https://github.com/wmfs/sharepoint/commit/76f30b09a00bdc5d681954e0384e7bf1f5cf173d))
* **deps-dev:** bump semantic-release from 15.14.0 to 17.0.4 ([3289448](https://github.com/wmfs/sharepoint/commit/3289448df06d3bb6267391938797e55f3e1ea1f0))
* **deps-dev:** bump semantic-release from 17.0.4 to 17.0.5 ([1e637eb](https://github.com/wmfs/sharepoint/commit/1e637eb3cff224ebbae611beeb9e575e42fd6360))
* **deps-dev:** bump semantic-release from 17.0.5 to 17.0.6 ([469031f](https://github.com/wmfs/sharepoint/commit/469031ff6e2d768ee0833db1fa0661f1e9d20fc7))
* **deps-dev:** bump semantic-release from 17.0.6 to 17.0.7 ([b0856a4](https://github.com/wmfs/sharepoint/commit/b0856a42b41037b246003f7c940855143b6a41f6))
* **deps-dev:** bump semantic-release from 17.0.7 to 17.0.8 ([676db29](https://github.com/wmfs/sharepoint/commit/676db2931f314a5bc814fb8d38527868ce29a0a8))
* **deps-dev:** bump standard from 14.3.1 to 14.3.2 ([56e43b0](https://github.com/wmfs/sharepoint/commit/56e43b07d0e78689732b80e4d42e6c22d1767ae0))
* **deps-dev:** bump standard from 14.3.2 to 14.3.3 ([acbe123](https://github.com/wmfs/sharepoint/commit/acbe12343d375c98c1162ac6a727b4fbd9fa61eb))
* **deps-dev:** bump standard from 14.3.3 to 14.3.4 ([b99016d](https://github.com/wmfs/sharepoint/commit/b99016dacb13e47cbd05e1b9bc3dc41199eb32cd))


### ⚙️ Continuous Integrations

* **circle:** add context env var config to config.yml ([7b96895](https://github.com/wmfs/sharepoint/commit/7b968958eba21f7b3edf162376da4cb9a22d1c76))

# [1.3.0](https://github.com/wmfs/sharepoint/compare/v1.2.0...v1.3.0) (2020-01-23)


### 🛠 Builds

* **deps:** bump axios from 0.19.1 to 0.19.2 ([992b9fa](https://github.com/wmfs/sharepoint/commit/992b9fa8f918db0137728573fedcee840e3a589d))
* **deps-dev:** bump conventional-changelog-metahub from 3.0.0 to 4.0.0 ([02b66c5](https://github.com/wmfs/sharepoint/commit/02b66c58d5008a88b6d2df7462bdee31fc55b047))

# [1.2.0](https://github.com/wmfs/sharepoint/compare/v1.1.0...v1.2.0) (2020-01-08)


### 🛠 Builds

* **deps:** bump axios from 0.19.0 to 0.19.1 ([#52](https://github.com/wmfs/sharepoint/issues/52)) ([7dfcf94](https://github.com/wmfs/sharepoint/commit/7dfcf94a36f38b7bd86ac02aba4b3b6f0a9eb8d4))
* **deps-dev:** bump packages ([fb40b30](https://github.com/wmfs/sharepoint/commit/fb40b30bf8809db2a138ffcef879ffe66f5ed673))
* **deps-dev:** update [@semantic-release](https://github.com/semantic-release)/git requirement ([3773b14](https://github.com/wmfs/sharepoint/commit/3773b14021c29c4bcbc78592a517cc7cd72ea0ca))
* **deps-dev:** update dev dependancies ([0928f37](https://github.com/wmfs/sharepoint/commit/0928f37b0b05d6f7937ff77267e23783d371cad8))
* **deps-dev:** update semantic-release requirement ([9711a8b](https://github.com/wmfs/sharepoint/commit/9711a8be733f7db6f235019296f72804e5333668))


### 📚 Documentation

* Add CircleCI badge ([a448924](https://github.com/wmfs/sharepoint/commit/a4489245df53878b5e550c2641b207ce56d368de))


### ⚙️ Continuous Integrations

* **circle:** Add CircleCI config ([f3595ae](https://github.com/wmfs/sharepoint/commit/f3595ae799d0c220a83c71bf96719c3755b10591))
* **travis:** Remove Travis config ([4f2713f](https://github.com/wmfs/sharepoint/commit/4f2713f43d7232aa57a6385fe0685ba146c61e4e))


### ♻️ Chores

* **deps-dev:** bump mocha from 6.2.2 to 7.0.0 ([dff38c3](https://github.com/wmfs/sharepoint/commit/dff38c3580a1aa6633a56c041a712bfd9579adb0))
* **deps-dev:** update codecov requirement from 3.1.0 to 3.5.0 ([ef08c90](https://github.com/wmfs/sharepoint/commit/ef08c902f532e7dc74cb0c78c0e2417124c958a5))
* **deps-dev:** update mocha requirement from 5.2.0 to 6.1.4 ([5b06603](https://github.com/wmfs/sharepoint/commit/5b066030cf1f4fe91bb4d6a9a64987d39c345b15))
* **deps-dev:** update nyc requirement from 13.1.0 to 14.1.1 ([7b290c6](https://github.com/wmfs/sharepoint/commit/7b290c67cfa29913a7458352cb9f18634d82ad82))
* **deps-dev:** update standard requirement from 12.0.1 to 14.3.1 ([ab3997b](https://github.com/wmfs/sharepoint/commit/ab3997b9131ed1d5d4f51ddec1eefa2a54106b31))

# [1.1.0](https://github.com/wmfs/sharepoint/compare/v1.0.1...v1.1.0) (2019-06-10)


### 🛠 Builds

* **deps:** update axios requirement from 0.18.0 to 0.19.0 ([#8](https://github.com/wmfs/sharepoint/issues/8)) ([141e093](https://github.com/wmfs/sharepoint/commit/141e093))


### 📚 Documentation

* update README.md ([ed226c0](https://github.com/wmfs/sharepoint/commit/ed226c0))
* update README.md (typo) ([8ac2f70](https://github.com/wmfs/sharepoint/commit/8ac2f70))

## [1.0.1](https://github.com/wmfs/sharepoint/compare/v1.0.0...v1.0.1) (2019-01-04)


### 🐛 Bug Fixes

* simplify way paths are passed in for creating folders etc ([4da514c](https://github.com/wmfs/sharepoint/commit/4da514c))


### 📚 Documentation

* remove coverage badge due to skipping tests ([52dc591](https://github.com/wmfs/sharepoint/commit/52dc591))

# 1.0.0 (2019-01-03)


### 🐛 Bug Fixes

* attempt release...? ([1a8fc5f](https://github.com/wmfs/sharepoint/commit/1a8fc5f))


### 📦 Code Refactoring

* get form digest value on each function ([ce86d5c](https://github.com/wmfs/sharepoint/commit/ce86d5c))


### 📚 Documentation

* added badges ([137a03d](https://github.com/wmfs/sharepoint/commit/137a03d))
* added badges ([2a00058](https://github.com/wmfs/sharepoint/commit/2a00058))
* added badges to README.md ([a3073f8](https://github.com/wmfs/sharepoint/commit/a3073f8))


### 🚨 Tests

* improve authenticate testing ([4e57d9e](https://github.com/wmfs/sharepoint/commit/4e57d9e))
* update form digest value test ([55f52d0](https://github.com/wmfs/sharepoint/commit/55f52d0))


### ⚙️ Continuous Integrations

* add standard to test script ([d564077](https://github.com/wmfs/sharepoint/commit/d564077))
* set up semantic release/travis config ([9f4e257](https://github.com/wmfs/sharepoint/commit/9f4e257))


### ♻️ Chores

* update deps ([b61439d](https://github.com/wmfs/sharepoint/commit/b61439d))


### 💎 Styles

* standard fix ([c381009](https://github.com/wmfs/sharepoint/commit/c381009))
