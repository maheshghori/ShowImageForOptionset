# ShowImageForOptionset
Some of the commands to remember for building project and creating Dynamics 365 Solution
# Open Developer Command Prompt for VS 2017/2019 and use below commands at different stages

mkdir RemainingCharacters
cd D:\PCF\ShowImageForOptionset
pac pcf init --namespace PCFControlGallery --name ShowImageForOptionset --template field
npm install
code .
npm run build
npm start
mkdir SolutionPackage
cd SolutionPackage
pac solution init --publisher-name dev --publisher-prefix dev
pac solution add-reference --path D:\PCF\ShowImageForOptionset [path to pcfproj file]

MSBUILD /t:restore
MSBUILD


