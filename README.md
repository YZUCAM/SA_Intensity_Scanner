<div align="center">
<table>
  <tr>
    <td rowspan="2">
      <img src="https://raw.githubusercontent.com/YZUCAM/SA_Intensity_Scanner/main/docsrc/I_Scan_UI.png" width="400"/>
    </td>
    <td>
      <img src="https://raw.githubusercontent.com/YZUCAM/SA_Intensity_Scanner/main/docsrc/ophir.png" width="200"/>
    </td>
  </tr>
  <tr>
    <td>
      <img src="https://raw.githubusercontent.com/YZUCAM/SA_Intensity_Scanner/main/docsrc/rotation_stage.png" width="200"/>
    </td>
  </tr>
</table>
</div>

# SA_Intensity_Scanner
A CLI based graphene saturable absorber intensity scanning program.

The program is used to synchronize the rotation stage and data acquisition. There is a halfwave plate installed in rotation stage. The different rotation angle will control the different incident laser power to the absorber material. Ophir records the transmitted laser power.

## Installation
conda install conda-forge::thorlabs-apt-protocol<br>
conda install anaconda::pywin32

